Option Explicit On
Option Strict Off

Public Class FrmLaporanTrialBalance
    Public UserCode As String = ""
    Public BackColorAll As Color = Color.LightSeaGreen

    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjHakAkses As New Core_Module.ClsHakAkses
    Dim DTA_Beginning_Balance As New DseLaporanTrialBalanceTableAdapters.tbl_beginning_balanceTableAdapter
    Dim TBL_Beginning_Balance As New DseLaporanTrialBalance.tbl_beginning_balanceDataTable
    Dim DTA_Jurnal_Trial_Balance As New DseLaporanTrialBalanceTableAdapters.jurnal_trial_balanceTableAdapter
    Dim TBL_Jurnal_Trial_Balance As New DseLaporanTrialBalance.jurnal_trial_balanceDataTable
    Dim DTA_Jurnal_Adjustment As New DseLaporanTrialBalanceTableAdapters.jurnal_adjustmentTableAdapter
    Dim TBL_Jurnal_Adjustment As New DseLaporanTrialBalance.jurnal_adjustmentDataTable
    Dim DTA_TrialBalance As New DseLaporanTrialBalanceTableAdapters.Trial_BalanceTableAdapter
    Dim ObjJurnal As New ClsTransaksiJournal
    Dim ObjLicense As New Core_Module.ClsLicense

    Dim report As New CrpLaporanTrialBalance
    Dim report_comparison As New CrpLaporanTrialBalance_Comparison
    Dim report_comparison2 As New CrpLaporanTrialBalance_Comparison2
    Dim DS As New DseLaporanTrialBalance
    Dim ObjBudget As New ClsMasterBudget
    Dim Def_Prec As Integer = 0
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim Negatif As Integer = 0

    Private Sub FrmLaporanTrialBalance_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.Def_Prec = Me.ObjCur.GetDecimal_DefaultCurrency
            Me.cboFormatLaporan.SelectedIndex = 0
            Me.cboJenis.SelectedIndex = 1 '--> selalu adjusted

            Me.dtpFrom_I.Value = Me.ObjCD.GetFirstDateofTheMonth(Today)
            Me.dtpFrom_II.Value = Me.ObjCD.GetFirstDateofTheMonth(Today)
            Me.dtpTo_I.Value = Me.ObjCD.GetFirstDateofTheMonth(Today)
            Me.dtpTo_II.Value = Me.ObjCD.GetFirstDateofTheMonth(Today)

            'jika tidak ada hak akses atas job, maka kolom job ditutup
            'If Me.ObjHakAkses.ApakahSudahAda("FrmMasterJob") = False Then
            Me.LblJob.Visible = False
            Me.TxtJobCode.Visible = False
            Me.TxtJobCode.TabStop = False
            Me.btnFindJob.Visible = False
            'End If

            'jika tidak ada hak akses atas dept, maka kolom dept ditutup
            'If Me.ObjHakAkses.ApakahSudahAda("FrmMasterDepartement") = False Then
            Me.LblDept.Visible = False
            Me.TxtDeptCode.Visible = False
            Me.TxtDeptCode.TabStop = False
            Me.btnFindDept.Visible = False
            'End If

            Me.cboFormatLaporan.Items.RemoveAt(3)
            Me.MinimizeBox = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboFormatLaporan_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFormatLaporan.SelectedIndexChanged
        Try
            If Me.cboFormatLaporan.SelectedIndex = 1 Or Me.cboFormatLaporan.SelectedIndex = 2 Then
                Me.LblPeriodII.Enabled = True
                Me.dtpTo_II.Enabled = True
            Else
                Me.LblPeriodII.Enabled = False
                Me.dtpTo_II.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiTabel(ByVal Adjusted As Boolean, ByVal vStartDate As Date, ByVal vEndDate As Date, _
    ByVal TBL As DseLaporanTrialBalance.Trial_BalanceDataTable)
        Dim BalanceDate As Date = Me.ObjCD.GetFirstDateofTheMonth(vStartDate) - TimeSpan.FromDays(1)
        vStartDate = Me.ObjCD.GetFirstDateofTheMonth(vStartDate)
        vEndDate = Me.ObjCD.GetLasDateOfTheMonth(vEndDate)
        Dim row As DseLaporanTrialBalance.Trial_BalanceRow
        Dim rowBeginningBalance As DseLaporanTrialBalance.tbl_beginning_balanceRow
        Dim row_TrialBalance As DseLaporanTrialBalance.jurnal_trial_balanceRow
        Dim row_Adjustment As DseLaporanTrialBalance.jurnal_adjustmentRow
        Dim filter As String = ""
        Dim rowView As DataRowView
        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim PrimaryAkunCYE As Double = -1
        Dim CYE As Decimal = 0
        Dim TipePerkiraan1 As ClsMasterAkun.EnumTipeAkun = ClsMasterAkun.EnumTipeAkun.Cost_of_Goods_Sales
        Dim Hapus As Boolean = True

        Try
            'untuk cetak trial balance normal, yang ditampilkan adalah nilai: BB + SL + PR
            'maksudnya:
            '1. Cari semua akun yang ada dengan begining balance pada awal period (kalo start period Jul 2010 --> begining balance
            '   adalah saldo buku besar pada 30 jun 2010
            'If Me.TxtPrimaryDept.Text = "-1" And Me.TxtPrimaryJob.Text = "-1" Then
            '    Me.DTA_Beginning_Balance.FillBy_All(Me.TBL_Beginning_Balance, BalanceDate)
            'ElseIf Me.TxtPrimaryDept.Text <> "-1" And Me.TxtPrimaryJob.Text = "-1" Then
            '    Me.DTA_Beginning_Balance.FillBy_PrimaryDept(Me.TBL_Beginning_Balance, BalanceDate, Me.TxtPrimaryDept.Text)
            'ElseIf Me.TxtPrimaryDept.Text = "-1" And Me.TxtPrimaryJob.Text <> "-1" Then
            '    Me.DTA_Beginning_Balance.FillBy_PrimaryJob(Me.TBL_Beginning_Balance, BalanceDate, Me.TxtPrimaryJob.Text)
            'Else
            Me.DTA_Beginning_Balance.Fill(Me.TBL_Beginning_Balance, BalanceDate)
            'End If

            '2. Tambahkan nilai tersebut dengan nilai jurnal yang terjadi selama periode yang diminta (From To --> jun - jul = 1 jun - 31 juli)
            '   kecuali jurnal adjustment dan jurnal closing'
            If Me.TxtPrimaryDept.Text = "-1" And Me.TxtPrimaryJob.Text = "-1" Then
                Me.DTA_Jurnal_Trial_Balance.FillBy_All(Me.TBL_Jurnal_Trial_Balance, vEndDate, vStartDate)
            ElseIf Me.TxtPrimaryDept.Text <> "-1" And Me.TxtPrimaryJob.Text = "-1" Then
                Me.DTA_Jurnal_Trial_Balance.FillBy_PrimaryDept(Me.TBL_Jurnal_Trial_Balance, vEndDate, vStartDate, Me.TxtPrimaryDept.Text)
            ElseIf Me.TxtPrimaryDept.Text = "-1" And Me.TxtPrimaryJob.Text <> "-1" Then
                Me.DTA_Jurnal_Trial_Balance.FillBy_PrimaryJob(Me.TBL_Jurnal_Trial_Balance, vEndDate, vStartDate, Me.TxtPrimaryJob.Text)
            Else
                Me.DTA_Jurnal_Trial_Balance.Fill(Me.TBL_Jurnal_Trial_Balance, vEndDate, vStartDate, Me.TxtPrimaryDept.Text, Me.TxtPrimaryJob.Text)
            End If

            '3. Untuk adjusted tambahkan nilainya dengan jurnal closing --> penambahan harus memperhitungkan posisi saldo normal, kalo saldo
            '   normal 1 maka posisi saldo normal di debet dan kumpulkan semua nilai didebet, hal berkebalikan berlaku juga.
            If Adjusted = True Then
                If Me.TxtPrimaryDept.Text = "-1" And Me.TxtPrimaryJob.Text = "-1" Then
                    Me.DTA_Jurnal_Adjustment.FillBy_All(Me.TBL_Jurnal_Adjustment, vEndDate, vStartDate)
                ElseIf Me.TxtPrimaryDept.Text <> "-1" And Me.TxtPrimaryJob.Text = "-1" Then
                    Me.DTA_Jurnal_Adjustment.FillBy_PrimaryDept(Me.TBL_Jurnal_Adjustment, vEndDate, vStartDate, Me.TxtPrimaryDept.Text)
                ElseIf Me.TxtPrimaryDept.Text = "-1" And Me.TxtPrimaryJob.Text <> "-1" Then
                    Me.DTA_Jurnal_Adjustment.FillBy_PrimaryJob(Me.TBL_Jurnal_Adjustment, vEndDate, vStartDate, Me.TxtPrimaryJob.Text)
                Else
                    Me.DTA_Jurnal_Adjustment.Fill(Me.TBL_Jurnal_Adjustment, vEndDate, vStartDate, Me.TxtPrimaryDept.Text, Me.TxtPrimaryJob.Text)
                End If
            Else
                Me.TBL_Jurnal_Adjustment.Rows.Clear()
            End If
            'penjumlahan beginning balance dengan jurnal yang ada
            For Each rowBeginningBalance In Me.TBL_Beginning_Balance.Rows
                filter = "Primary =" + rowBeginningBalance.Primary.ToString
                Me.TBL_Jurnal_Trial_Balance.DefaultView.RowFilter = filter
                For Each rowView In Me.TBL_Jurnal_Trial_Balance.DefaultView
                    row_TrialBalance = rowView.Row
                    rowBeginningBalance.Debet = rowBeginningBalance.Debet + row_TrialBalance.Debet
                    rowBeginningBalance.Kredit = rowBeginningBalance.Kredit + row_TrialBalance.Kredit
                Next

                If Adjusted = True Then
                    Me.TBL_Jurnal_Adjustment.DefaultView.RowFilter = filter
                    For Each rowView In Me.TBL_Jurnal_Adjustment.DefaultView
                        row_Adjustment = rowView.Row
                        rowBeginningBalance.Debet = rowBeginningBalance.Debet + row_Adjustment.Debet
                        rowBeginningBalance.Kredit = rowBeginningBalance.Kredit + row_Adjustment.Kredit
                    Next
                End If
            Next

            '4. Pindahkan ke tabel trial balance
            Me.DTA_TrialBalance.Fill(TBL)
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
            Next

            '4.a. Tambahkan akun current year earning, jika pada bulan tersebut tidak ada jurnal closing
            If Me.ObjJurnal.GetJml(ClsTransaksiJournal.EnumTipeJournal.ClosingJournal, _
            Me.ObjCD.GetFisrtDateOfTheYear(vEndDate), Me.ObjCD.GetLasDateOfTheMonth(vEndDate)) = 0 Then
                CYE = Me.ObjJurnal.GetCurrentYearEarning1(vEndDate, -1, -1)
                PrimaryAkunCYE = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_CurrenctYearEarning)

                For Each row In TBL.Rows
                    If row.Primary_Akun = PrimaryAkunCYE Then
                        If row.Saldo_Normal = 0 Then
                            row.Debet = row.Debet + CYE
                            row.Credit = 0
                        Else
                            row.Credit = row.Credit + CYE
                            row.Debet = 0
                        End If
                    End If
                Next
            End If

            '5. Hitung sum untuk setiap data header --> kalo mentok pada bagian ini langsung hub gw, kemungkinan ini pake rekursif
            Me.UpdateSaldoHeader(TBL)

            '6. Use total
            If Me.chkUseTotal.Checked = True Then
                i = 0
                While i < TBL.Rows.Count
                    row = TBL.Rows(i)
                    If row.Header = 1 Then
                        If row.IsTotalNull = True Then
                            Me.InsertTotal(i, row.Level_Perkiraan, TBL)
                            row.Total = 0
                        End If
                    Else
                        row.Total = 0
                    End If
                    i += 1
                End While
            Else
                For Each row In TBL.Rows
                    row.Total = -1
                Next
            End If
            '7. Hide data 
            If Me.chkShowLvl2.Checked = False Then
                i = 2
            ElseIf Me.chkShowLvl3.Checked = False Then
                i = 3
            ElseIf Me.chkShowLvl4.Checked = False Then
                i = 4
            Else
                i = 5
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

            'jika semua akun cogs bernilai 0, maka buang akun cogs
            Hapus = True
            For Each row In TBL.Rows
                If row.Tipe_Perkiraan = CInt(TipePerkiraan1) Then
                    If row.Debet <> 0 Or row.Credit <> 0 Then
                        Hapus = False
                    End If
                End If
            Next
            If Hapus = True Then
                i = 0
                While i <= TBL.Rows.Count - 1
                    row = TBL.Rows(i)
                    If row.Tipe_Perkiraan = CInt(TipePerkiraan1) Then
                        TBL.Rows.Remove(row)
                    Else
                        i += 1
                    End If
                End While
            End If

            Hapus = True
            TipePerkiraan1 = ClsMasterAkun.EnumTipeAkun.Tax
            For Each row In TBL.Rows
                If row.Tipe_Perkiraan = CInt(TipePerkiraan1) Then
                    If row.Debet <> 0 Or row.Credit <> 0 Then
                        Hapus = False
                    End If
                End If
            Next
            If Hapus = True Then
                i = 0
                While i <= TBL.Rows.Count - 1
                    row = TBL.Rows(i)
                    If row.Tipe_Perkiraan = CInt(TipePerkiraan1) Then
                        TBL.Rows.Remove(row)
                    Else
                        i += 1
                    End If
                End While
            End If
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
                    .Akun = "Total " + row.Description
                ElseIf row.Level_Perkiraan = 2 Then
                    .Akun = space + "Total " + row.Description
                ElseIf row.Level_Perkiraan = 3 Then
                    .Akun = space + space + "Total " + row.Description
                ElseIf row.Level_Perkiraan = 4 Then
                    .Akun = space + space + space + "Total " + row.Description
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
                .Saldo_Normal = row.Saldo_Normal
                .Description = row.Description
                .Header = row.Header
                .Level_Perkiraan = row.Level_Perkiraan
                .Primary_Akun = Me.GetPrimaryAkunMaks(vTBL) + 1
                .Saldo = row.Saldo
                .Total = 1
                .Tipe_Perkiraan = row.Tipe_Perkiraan
                .Urutan = row.Tipe_Perkiraan
                .Account_Code = row.Account_Code
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
    ByVal vTBL As DseLaporanTrialBalance.Trial_BalanceDataTable) As Decimal
        Dim TargetIndex, CurrentIndex As Integer
        Dim row As DseLaporanTrialBalance.Trial_BalanceRow
        Dim RetVal As Decimal = 0

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

    Private Sub CetakTrialBalance(ByVal Adjusted As Boolean)
        Dim TBL As DseLaporanTrialBalance.Trial_BalanceDataTable = Me.DS.Trial_Balance
        Dim row As DseLaporanTrialBalance.Trial_BalanceRow
        Dim rowCetak As DseLaporanTrialBalance.TblCetakRow
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim i As Double = 0
        Dim vDtpFrom_1 As Date = Now
        Dim vDtpTo_1 As Date = Now
        Dim splitakun(1) As String

        Try
            'untuk cetak trial balance normal, yang ditampilkan adalah nilai: BB + SL + PR
            'maksudnya:
            '1. Isikan tabel dengan fungsi yang sudah disediakan
            Me.IsiTabel(Adjusted, Me.dtpFrom_I.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpTo_I.Value), TBL)
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
                    .Total = row.Total
                    .Aktif = row.Aktif
                    .Saldo_Normal = row.Saldo_Normal
                    .Level_Perkiraan = row.Level_Perkiraan
                    .Bold = row.Bold
                    .Urutan = row.Tipe_Perkiraan

                    ''TR//Perubahan posisi akun menjadi rata kiri sama misahin No / Name Account
                    If Me.chkShowAccountCode.Checked = False Then
                        .Akun = row.Akun.Trim
                    Else
                        If row.Total = 1 Then
                            .Akun_Name = row.Akun.Trim
                        Else
                            .Akun_No = row.Account_Code
                            .Akun_Name = row.Description
                        End If
                    End If
                    
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

            vDtpFrom_1 = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_I.Value)
            vDtpTo_1 = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpTo_I.Value)

            '''TR//Perubahan Tampilan Akun menjadi rata kiri semua & misahin antara no Akun dan namannya
            'For x As Integer = 0 To DS.TblCetak.Rows.Count - 1
            '    rowCetak = DS.TblCetak.Rows(x)
            '    rowCetak.Akun = rowCetak.Akun.Trim
            '    If Me.chkShowAccountCode.Checked = True Then
            '        splitakun = Split(rowCetak.Akun, "-")
            '        If splitakun.Length <> 1 Then
            '            If splitakun.Length = 2 Then
            '                rowCetak.Akun_No = splitakun(0).Trim
            '                rowCetak.Akun_Name = splitakun(1).Trim
            '            ElseIf splitakun.Length > 2 Then
            '                For counter As Integer = 0 To splitakun.Length - 1
            '                    If IsNumeric(splitakun(counter).Trim) = True Then
            '                        If rowCetak.Akun_Name = "" Then
            '                            If rowCetak.Akun_No = "" Then
            '                                rowCetak.Akun_No = splitakun(counter).Trim
            '                            Else
            '                                rowCetak.Akun_No &= "-" & splitakun(counter).Trim
            '                            End If
            '                        Else
            '                            rowCetak.Akun_Name &= "-" & splitakun(counter)
            '                        End If
            '                    Else
            '                        If rowCetak.Akun_Name = "" Then
            '                            rowCetak.Akun_Name = splitakun(counter).TrimStart
            '                        Else
            '                            rowCetak.Akun_Name &= "-" & splitakun(counter)
            '                        End If
            '                    End If
            '                Next
            '            End If
            '        Else
            '            rowCetak.Akun_Name = splitakun(0).Trim
            '        End If

            '    End If
            'Next

            Me.report.SetDataSource(DS)
            Me.report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, Me.Def_Prec)
            Me.report.SetParameterValue(report.Parameter_TanggalDari.ParameterFieldName, vDtpFrom_1)
            Me.report.SetParameterValue(report.Parameter_TanggalSampai.ParameterFieldName, vDtpTo_1)
            If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                Me.Negatif = 1
            Else
                Me.Negatif = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
            End If
            Me.report.SetParameterValue(report.Parameter_Negatif.ParameterFieldName, Me.Negatif)
            Me.report.SetParameterValue(report.Parameter_Counter.ParameterFieldName, 0)
            If Me.chkShowAccountCode.Checked = True Then
                Me.report.SetParameterValue(report.Parameter_ShowNoAccount.ParameterFieldName, 1)
            Else
                Me.report.SetParameterValue(report.Parameter_ShowNoAccount.ParameterFieldName, 0)
            End If
            Me.crv1.ReportSource = Me.report
            Me.crv1.Refresh()
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
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

    Private Sub CetakTrialBalance_Comparison_Period(ByVal Adjusted As Boolean)
        Dim TBL As DseLaporanTrialBalance.Trial_BalanceDataTable = Me.DS.Trial_Balance
        Dim TBL2 As New DseLaporanTrialBalance.Trial_BalanceDataTable
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
        Dim splitAkun(1) As String

        Try
            'untuk cetak trial balance normal, yang ditampilkan adalah nilai: BB + SL + PR
            'maksudnya:
            '1. Isikan tabel dengan fungsi yang sudah disediakan
            Me.IsiTabel(Adjusted, Me.dtpFrom_I.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpTo_I.Value), TBL)
            Me.IsiTabel(Adjusted, Me.dtpFrom_II.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpTo_II.Value), TBL2)
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
                    ''TR//Perubahan posisi akun menjadi rata kiri sama misahin No / Name Account
                    If Me.chkShowAccountCode.Checked = False Then
                        .Akun = row.Akun.Trim
                    Else
                        If row.Total = 1 Then
                            .Akun_Name = row.Akun.Trim
                        Else
                            .Akun_No = row.Account_Code
                            .Akun_Name = row.Description
                        End If
                    End If

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
                    'hitung variance
                    .Variance = .Saldo_Period_II - .Saldo_Period_I
                    If .Saldo_Period_I <> 0 Then
                        .Variance_Persen = Math.Round(((.Saldo_Period_II - .Saldo_Period_I) / _
                        .Saldo_Period_I) * 100, 2, MidpointRounding.AwayFromZero)
                    Else
                        .Variance_Persen = 0
                    End If
                End With

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
            vDtpTo_1 = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpTo_I.Value)
            vDtpFrom_2 = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_II.Value)
            vDtpTo_2 = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpTo_II.Value)


            ''TR//Perubahan posisi pada akun, dibuat sama rata kiri. Sama misahin antara No / Name account
            'If Me.cboFormatLaporan.SelectedIndex = 1 Then
            '    For x As Integer = 0 To DS.TblCetak.Rows.Count - 1
            '        rowCetak = DS.TblCetak.Rows(x)
            '        rowCetak.Akun = rowCetak.Akun.Trim
            '        If Me.chkShowAccountCode.Checked = True Then
            '            splitAkun = Split(rowCetak.Akun, "-")
            '            If splitAkun.Length <> 1 Then
            '                rowCetak.Akun_No = splitAkun(0).Trim
            '                rowCetak.Akun_Name = splitAkun(1).Trim
            '            Else
            '                rowCetak.Akun_Name = splitAkun(0).Trim
            '            End If

            '        End If
            '    Next
            'End If
      
         

            If Me.cboFormatLaporan.SelectedIndex = 1 Then
                Me.report_comparison.SetDataSource(DS)
                Me.report_comparison.SetParameterValue(report_comparison.Parameter_FormatDecimal.ParameterFieldName, Me.Def_Prec)
                Me.report_comparison.SetParameterValue(report_comparison.Parameter_TanggalDari_1.ParameterFieldName, vDtpFrom_1)
                Me.report_comparison.SetParameterValue(report_comparison.Parameter_TanggalSampai_1.ParameterFieldName, vDtpTo_1)
                Me.report_comparison.SetParameterValue(report_comparison.Parameter_TanggalDari_2.ParameterFieldName, vDtpFrom_2)
                Me.report_comparison.SetParameterValue(report_comparison.Parameter_TanggalSampai_2.ParameterFieldName, vDtpTo_2)
                Me.report_comparison.SetParameterValue(report_comparison.Parameter_Tipe.ParameterFieldName, tipe)
                If Me.chkShowAccountCode.Checked = True Then
                    Me.report_comparison.SetParameterValue(report_comparison.Parameter_ShowAccount.ParameterFieldName, 1)
                Else
                    Me.report_comparison.SetParameterValue(report_comparison.Parameter_ShowAccount.ParameterFieldName, 0)
                End If
                If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                    Me.Negatif = 1
                Else
                    Me.Negatif = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
                End If
                Me.report_comparison.SetParameterValue(report_comparison.Parameter_Negatif.ParameterFieldName, Me.Negatif)

                Me.crv1.ReportSource = Me.report_comparison
                Me.crv1.Refresh()
                Me.crv1.Show()
            Else
                Me.report_comparison2.SetDataSource(DS)
                Me.report_comparison2.SetParameterValue(report_comparison2.Parameter_FormatDecimal.ParameterFieldName, Me.Def_Prec)
                Me.report_comparison2.SetParameterValue(report_comparison2.Parameter_TanggalSampai.ParameterFieldName, Me.dtpTo_I.Value)
                Me.report_comparison2.SetParameterValue(report_comparison2.Parameter_TanggalSampaiII.ParameterFieldName, Me.dtpTo_II.Value)
                If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                    Me.Negatif = 1
                Else
                    Me.Negatif = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
                End If
                Me.report_comparison2.SetParameterValue(report_comparison2.Parameter_Negatif.ParameterFieldName, Me.Negatif)

                Me.crv1.ReportSource = Me.report_comparison2
                Me.crv1.Refresh()
                Me.crv1.Show()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
            TBL2.Dispose()
        End Try
    End Sub

    Private Sub CetakTrialBalance_Comparison_Budget(ByVal Adjusted As Boolean)
        Dim TBL As DseLaporanTrialBalance.Trial_BalanceDataTable = Me.DS.Trial_Balance
        Dim TBL2 As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim row As DseLaporanTrialBalance.Trial_BalanceRow
        Dim rowCetak As DseLaporanTrialBalance.TblCetakRow
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim i As Double = 0
        Dim filter As String = ""
        Dim vDtpFrom_1 As Date = Now
        Dim vDtpTo_1 As Date = Now
        Dim vDtpFrom_2 As Date = Now
        Dim vDtpTo_2 As Date = Now
        Dim tipe As Double = 1
        Dim SplitAkun(1) As String
        Try
            'untuk cetak trial balance normal, yang ditampilkan adalah nilai: BB + SL + PR
            'maksudnya:
            '1. Isikan tabel dengan fungsi yang sudah disediakan
            Me.IsiTabel(Adjusted, Me.dtpFrom_I.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpTo_I.Value), TBL)
            MsgBox("Logic error, this function should not be executed. Please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
            'Me.ObjBudget.GetBudget(Me.ObjCD.GetLasDateOfTheMonth(Me.dtpTo_I.Value), TBL2)
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

                    ''TR//Perubahan posisi akun menjadi rata kiri sama misahin No / Name Account
                    If Me.chkShowAccountCode.Checked = False Then
                        .Akun = row.Akun.Trim
                    Else
                        If row.Total = 1 Then
                            .Akun_Name = row.Akun.Trim
                        Else
                            .Akun_No = row.Account_Code
                            .Akun_Name = row.Description
                        End If
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
            vDtpTo_1 = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpTo_I.Value)
            vDtpFrom_2 = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom_II.Value)
            vDtpTo_2 = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpTo_II.Value)

            ''TR//Perubahan posisi akun menjadi rata kiri sama misahin No / Name Account
            'For x As Integer = 0 To DS.TblCetak.Rows.Count - 1
            '    rowCetak = DS.TblCetak.Rows(x)
            '    rowCetak.Akun = rowCetak.Akun.Trim
            '    If Me.chkShowAccountCode.Checked = True Then
            '        SplitAkun = Split(rowCetak.Akun, "-")
            '        If SplitAkun.Length <> 1 Then
            '            rowCetak.Akun_No = SplitAkun(0).Trim
            '            rowCetak.Akun_Name = SplitAkun(1).Trim
            '        Else
            '            rowCetak.Akun_Name = SplitAkun(0).Trim
            '        End If

            '    End If
            'Next

            Me.report_comparison.SetDataSource(DS)
            Me.report_comparison.SetParameterValue(report_comparison.Parameter_FormatDecimal.ParameterFieldName, Me.Def_Prec)
            Me.report_comparison.SetParameterValue(report_comparison.Parameter_TanggalDari_1.ParameterFieldName, vDtpFrom_1)
            Me.report_comparison.SetParameterValue(report_comparison.Parameter_TanggalSampai_1.ParameterFieldName, vDtpTo_1)
            Me.report_comparison.SetParameterValue(report_comparison.Parameter_TanggalDari_2.ParameterFieldName, vDtpFrom_2)
            Me.report_comparison.SetParameterValue(report_comparison.Parameter_TanggalSampai_2.ParameterFieldName, vDtpTo_2)
            Me.report_comparison.SetParameterValue(report_comparison.Parameter_Tipe.ParameterFieldName, tipe)

            If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                Me.Negatif = 1
            Else
                Me.Negatif = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
            End If
            Me.report_comparison.SetParameterValue(report_comparison.Parameter_Negatif.ParameterFieldName, Me.Negatif)

            If Me.chkShowAccountCode.Checked = True Then
                Me.report_comparison.SetParameterValue(report_comparison.Parameter_ShowAccount.ParameterFieldName, 1)
            Else
                Me.report_comparison.SetParameterValue(report_comparison.Parameter_ShowAccount.ParameterFieldName, 0)
            End If
            Me.crv1.ReportSource = Me.report_comparison
            Me.crv1.Refresh()
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
            TBL2.Dispose()
        End Try
    End Sub

    Private Sub btnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShow.Click
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
            If Me.dtpTo_I.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpTo_I.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.dtpTo_II.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpTo_II.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            If Me.cboJenis.SelectedIndex = 0 Then
            Else
                If Me.cboFormatLaporan.SelectedIndex = 0 Then
                    If Me.dtpFrom_I.Value.Date > Me.dtpTo_I.Value.Date Then
                        Me.dtpFrom_I.Value = Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpTo_I.Value)
                    End If
                    Me.CetakTrialBalance(True)
                ElseIf Me.cboFormatLaporan.SelectedIndex = 1 Or _
                Me.cboFormatLaporan.SelectedIndex = 2 Then
                    If Me.dtpFrom_I.Value.Date > Me.dtpTo_I.Value.Date Then
                        Me.dtpFrom_I.Value = Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpTo_I.Value)
                    End If
                    If Me.dtpFrom_II.Value.Date > Me.dtpTo_II.Value.Date Then
                        Me.dtpFrom_II.Value = Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpTo_II.Value)
                    End If
                    Me.CetakTrialBalance_Comparison_Period(True)
                Else
                    If Me.dtpFrom_I.Value.Date > Me.dtpTo_I.Value.Date Then
                        Me.dtpFrom_I.Value = Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpTo_I.Value)
                    End If
                    Me.CetakTrialBalance_Comparison_Budget(True)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
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
                Me.chkShowLvl4.Enabled = True
            Else
                Me.chkShowLvl3.Enabled = False
                Me.chkShowLvl4.Enabled = False
                Me.chkShowLvl3.Checked = False
                Me.chkShowLvl4.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkShowLvl3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowLvl3.CheckedChanged
        Try
            If Me.chkShowLvl3.Checked = True Then
                Me.chkShowLvl4.Enabled = True
            Else
                Me.chkShowLvl4.Enabled = False
                Me.chkShowLvl4.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpTo_I_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpTo_I.Leave
        Me.BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Threading.Thread.Sleep(100)
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.dtpFrom_I.Value = Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpTo_I.Value)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpTo_II_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpTo_II.Leave
        Me.BackgroundWorker2.RunWorkerAsync()
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Threading.Thread.Sleep(100)
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            Me.dtpFrom_II.Value = Me.ObjCD.GetFisrtDateOfTheYear(Me.dtpTo_II.Value)
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
End Class