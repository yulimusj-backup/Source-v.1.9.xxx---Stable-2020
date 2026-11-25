Option Explicit On
Option Strict Off

Public Class FrmLaporan_ProfitNLoss
    Public Usercode As String = ""
    Public BackColorAll As Color = Color.LightSeaGreen

    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjClosingAcc As New AccModule.ClsTransaksiClosing
    Dim ObjHakAkses As New Core_Module.ClsHakAkses
    Dim DTA_Beginning_Balance As New DseLaporanTrialBalanceTableAdapters.tbl_beginning_balanceTableAdapter
    Dim TBL_Beginning_Balance As New DseLaporanTrialBalance.tbl_beginning_balanceDataTable
    Dim DTA_Jurnal_Trial_Balance As New DseLaporanTrialBalanceTableAdapters.jurnal_trial_balanceTableAdapter
    Dim TBL_Jurnal_Trial_Balance As New DseLaporanTrialBalance.jurnal_trial_balanceDataTable
    Dim DTA_Jurnal_Adjustment As New DseLaporanTrialBalanceTableAdapters.jurnal_adjustmentTableAdapter
    Dim TBL_Jurnal_Adjustment As New DseLaporanTrialBalance.jurnal_adjustmentDataTable
    Dim DTA_TrialBalance As New DseLaporanTrialBalanceTableAdapters.Trial_BalanceTableAdapter

    Dim report As New CrpLaporanProfitNLoss
    Dim report_comparison As New CrpLaporanProfitNLoss_Comparison
    Dim Report_6Months As New CrpLaporanProfitNLoss_6Months
    Dim Report_12Months As New CrpLaporanProfitNLoss_12Months
    Dim Report_QTR As New CrpLaporanProfitNLoss_12Months_QTR
    Dim reportChart As New CrpLaporanChart_IncomeStatement

    Dim DS As New DseLaporanTrialBalance
    Dim ObjBudget As New ClsMasterBudget

    Dim Def_Prec As Integer = 0
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim Negatif As Integer = 0
    Dim ConvertionDate As Date = #1/1/1950#
    Dim Obj_Closing As New ClsTransaksiClosing
    Dim ObjLicense As New Core_Module.ClsLicense

    Private Sub BuangKodeAkun(ByVal TBL As DseLaporanTrialBalance.Trial_BalanceDataTable)
        Dim row As DseLaporanTrialBalance.Trial_BalanceRow
        Dim Space As String = "     "
        Dim SpaceTotal As String = ""

        Try
            For Each row In TBL.Rows
                If row.Tipe_Perkiraan <> 9 And row.Tipe_Perkiraan <> 10 And row.Tipe_Perkiraan <> 11 Then
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
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanTrialBalance_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.Def_Prec = Me.ObjCur.GetDecimal_DefaultCurrency
            Me.cboFormatLaporan.SelectedIndex = 0

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

            Me.chkShowLvl2.Enabled = True
            Me.chkShowLvl3.Enabled = True
            Me.chkShowLvl4.Enabled = True
            Me.chkShowLvl5.Enabled = True
            Me.chkShowLvl6.Enabled = True
            Me.chkShowLvl7.Enabled = True
            Me.chkShowLvl8.Enabled = True

            'End If
            Me.ConvertionDate = Me.ObjClosingAcc.GetConvertionDate
            Me.chkUseTotal.Text = "Use Indentation"
            Me.MinimizeBox = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboFormatLaporan_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFormatLaporan.SelectedIndexChanged
        Try
            Me.gboPeriod_I.Text = "Period I"
            Me.LblPeriodFrom.Text = "From"
            Me.LblPeriodTo.Text = "To"
            Me.dtpTo_I.Visible = True
            Me.dtpTo_I.Enabled = True
            Me.dtpTo_I.TabStop = True
            Me.LblPeriodTo.Visible = True
            Me.LblPeriodTo.Enabled = True
            Me.dtpFrom_I.CustomFormat = "MMM yyyy"

            Me.chkShowCOGS.Checked = False
            Me.chkShowExpense.Checked = False
            Me.chkShowIncome.Checked = False
            Me.chkShowProfit.Checked = False
            Me.chkShowCOGS.TabStop = False
            Me.chkShowExpense.TabStop = False
            Me.chkShowIncome.TabStop = False
            Me.chkShowProfit.TabStop = False
            'Me.chkShowCOGS.Visible = False
            'Me.chkShowExpense.Visible = False
            'Me.chkShowIncome.Visible = False
            'Me.chkShowProfit.Visible = False

            If Me.cboFormatLaporan.SelectedIndex = 1 Then
                '2 period + vs budget
                Me.gboPeriod_II.Enabled = True
                'Me.gboPeriod_II.Visible = True
                Me.chkShowCOGS.Enabled = False
                Me.chkShowExpense.Enabled = False
                Me.chkShowIncome.Enabled = False
                Me.chkShowProfit.Enabled = False
            ElseIf Me.cboFormatLaporan.SelectedIndex = 0 Or Me.cboFormatLaporan.SelectedIndex = 2 Then
                '1 period
                Me.gboPeriod_II.Enabled = False
                'Me.gboPeriod_II.Visible = False
                Me.chkShowCOGS.Enabled = False
                Me.chkShowExpense.Enabled = False
                Me.chkShowIncome.Enabled = False
                Me.chkShowProfit.Enabled = False
            ElseIf Me.cboFormatLaporan.SelectedIndex = 3 Then
                Me.gboPeriod_II.Enabled = False
                'Me.gboPeriod_II.Visible = False
                Me.gboPeriod_I.Text = "Period"
                Me.LblPeriodFrom.Text = "Period I"
                Me.LblPeriodTo.Text = "Period II"
                Me.chkShowCOGS.Enabled = False
                Me.chkShowExpense.Enabled = False
                Me.chkShowIncome.Enabled = False
                Me.chkShowProfit.Enabled = False
            ElseIf Me.cboFormatLaporan.SelectedIndex = 4 Or Me.cboFormatLaporan.SelectedIndex = 5 Then
                Me.gboPeriod_II.Enabled = False
                'Me.gboPeriod_II.Visible = False
                Me.gboPeriod_I.Text = "Period"
                Me.LblPeriodFrom.Text = "From"
                Me.LblPeriodTo.Text = "Period II"
                'Me.dtpTo_I.Visible = False
                Me.dtpTo_I.TabStop = False
                'Me.LblPeriodTo.Visible = False
                Me.LblPeriodTo.Enabled = False
                Me.chkShowCOGS.Enabled = False
                Me.chkShowExpense.Enabled = False
                Me.chkShowIncome.Enabled = False
                Me.chkShowProfit.Enabled = False
            ElseIf Me.cboFormatLaporan.SelectedIndex = 6 Or Me.cboFormatLaporan.SelectedIndex = 7 Or Me.cboFormatLaporan.SelectedIndex = 8 Then
                Me.gboPeriod_II.Enabled = False
                'Me.gboPeriod_II.Visible = False
                Me.gboPeriod_I.Text = "Period"
                Me.LblPeriodFrom.Text = "Year"
                Me.LblPeriodTo.Text = "Period II"
                Me.dtpFrom_I.CustomFormat = "yyyy"
                'Me.dtpTo_I.Visible = False
                Me.dtpTo_I.Enabled = False
                Me.dtpTo_I.TabStop = False
                'Me.LblPeriodTo.Visible = False
                Me.LblPeriodTo.Enabled = False
                Me.chkShowCOGS.Enabled = False
                Me.chkShowExpense.Enabled = False
                Me.chkShowIncome.Enabled = False
                Me.chkShowProfit.Enabled = False
            ElseIf Me.cboFormatLaporan.SelectedIndex = 9 Then
                '1 period
                Me.gboPeriod_II.Enabled = False
                'Me.gboPeriod_II.Visible = False
                'check list berkaitan dengan chart
                Me.chkShowCOGS.Checked = True
                Me.chkShowExpense.Checked = True
                Me.chkShowIncome.Checked = True
                Me.chkShowProfit.Checked = True
                Me.chkShowCOGS.TabStop = True
                Me.chkShowExpense.TabStop = True
                Me.chkShowIncome.TabStop = True
                Me.chkShowProfit.TabStop = True
                'Me.chkShowCOGS.Visible = True
                'Me.chkShowExpense.Visible = True
                'Me.chkShowIncome.Visible = True
                'Me.chkShowProfit.Visible = True
                Me.chkShowCOGS.Enabled = True
                Me.chkShowExpense.Enabled = True
                Me.chkShowIncome.Enabled = True
                Me.chkShowProfit.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiTabel(ByVal Adjusted As Boolean, ByVal vStartDate As Date, ByVal vEndDate As Date, _
    ByVal TBL As DseLaporanTrialBalance.Trial_BalanceDataTable, Optional ByVal BuangCOGSNol As Boolean = True)
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
        Dim TBLDummy As New DseLaporanTrialBalance.Trial_BalanceDataTable
        Dim TipePerkiraan1 As ClsMasterAkun.EnumTipeAkun = ClsMasterAkun.EnumTipeAkun.Cost_of_Goods_Sales
        Dim Hapus As Boolean = True
        Dim NolKanBeginingBalance As Boolean = False
        Dim Space As String = "   "
        Dim Dummy As Integer = 0

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

            For Each rowBeginningBalance In Me.TBL_Beginning_Balance.Rows
                With rowBeginningBalance
                    .Debet = 0
                    .Kredit = 0

                    'revisi iwan: 2018-03-29, mengikut data mst yang implementasi ditengah taun
                    If rowBeginningBalance.Tampilan.Trim.StartsWith("1") Or _
                    rowBeginningBalance.Tampilan.Trim.StartsWith("2") Or _
                    rowBeginningBalance.Tampilan.Trim.StartsWith("3") Then
                    Else
                        .Beginning_Balance = 0
                    End If
                End With
            Next

            '2. Tambahkan nilai tersebut dengan nilai jurnal yang terjadi selama periode yang diminta (From To --> jun - jul = 1 jun - 31 juli)
            '   kecuali jurnal adjustment dan jurnal closing
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
                If Me.TxtPrimaryDept.Text <> "-1" And Me.TxtPrimaryJob.Text <> "-1" Then
                    Me.DTA_Jurnal_Adjustment.Fill(Me.TBL_Jurnal_Adjustment, vEndDate, vStartDate, Me.TxtPrimaryDept.Text, Me.TxtPrimaryJob.Text)
                ElseIf Me.TxtPrimaryDept.Text <> "-1" And Me.TxtPrimaryJob.Text = "-1" Then
                    Me.DTA_Jurnal_Adjustment.FillBy_PrimaryDept(Me.TBL_Jurnal_Adjustment, vEndDate, vStartDate, Me.TxtPrimaryDept.Text)
                ElseIf Me.TxtPrimaryDept.Text = "-1" And Me.TxtPrimaryJob.Text <> "-1" Then
                    Me.DTA_Jurnal_Adjustment.FillBy_PrimaryJob(Me.TBL_Jurnal_Adjustment, vEndDate, vStartDate, Me.TxtPrimaryJob.Text)
                Else
                    Me.DTA_Jurnal_Adjustment.FillBy_All(Me.TBL_Jurnal_Adjustment, vEndDate, vStartDate)
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

            '4.a. Pindahkan ke tabel trial balance dummy
            Me.DTA_TrialBalance.Fill(TBLDummy)
            '4.b. Buang akun asset, liabilities, equity
            j = 0
            While j < TBL.Rows.Count
                row = TBL.Rows(j)
                If row.Tipe_Perkiraan <= 2 Then
                    TBL.Rows.Remove(row)
                Else
                    j += 1
                End If
            End While
            '4.c pindahkan data debet kredit
            If Me.ConvertionDate.Month = 1 Then
                If Format(vStartDate, "MMM yyyy") <> _
                        Format(CDate("#" + (12).ToString + "/" + Me.ConvertionDate.Day.ToString + "/" + (Me.ConvertionDate.Year - 1).ToString + "#"), "MMM yyyy") Then
                    NolKanBeginingBalance = True
                End If
            Else
                If Format(vStartDate, "MMM yyyy") <> _
                        Format(CDate("#" + (Me.ConvertionDate.Month - 1).ToString + "/" + Me.ConvertionDate.Day.ToString + "/" + Me.ConvertionDate.Year.ToString + "#"), "MMM yyyy") Then
                    NolKanBeginingBalance = True
                End If
            End If

            For Each row In TBLDummy.Rows
                filter = "Primary =" + row.Primary_Akun.ToString
                Me.TBL_Beginning_Balance.DefaultView.RowFilter = filter

                If row.Tipe_Perkiraan <> ClsMasterAkun.EnumTipeAkun.Asset And row.Tipe_Perkiraan <> ClsMasterAkun.EnumTipeAkun.Equity And _
                row.Tipe_Perkiraan <> ClsMasterAkun.EnumTipeAkun.Liabilities Then
                    'revisi iwan: 2018-03-29, mengikut data mst yang implementasi ditengah taun
                    If NolKanBeginingBalance = True Then
                        row.Beginning_Balance = 0
                    End If
                End If

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
            '4.d. format urutan akun pada tabel
            Me.PindahkanData_Profit_N_Loss(TBLDummy, TBL)

            '5. Hitung sum untuk setiap data header --> kalo mentok pada bagian ini langsung hub gw, kemungkinan ini pake rekursif
            Me.UpdateSaldoHeader(TBL)

            '6. Use total
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

            '7. Hide data 
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

            '9. jika semua akun cogs bernilai 0, maka buang akun cogs
            If BuangCOGSNol = True Then
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
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub PindahkanData_Profit_N_Loss(ByVal vTBLSumber As DseLaporanTrialBalance.Trial_BalanceDataTable, _
    ByVal vTBLTarget As DseLaporanTrialBalance.Trial_BalanceDataTable)
        Dim TotalIncome, TotalCOGS, TotalExpense, TotalOtherIncome, TotalOtherExpense, TotalTax As Double
        Dim rowBaru As DseLaporanTrialBalance.Trial_BalanceRow

        Try
            'kosongkan tabel target
            vTBLTarget.Rows.Clear()
            'start pindahkan data
            '1. pindahkan data income
            TotalIncome = Me.PindahkanData(vTBLSumber, vTBLTarget, ClsMasterAkun.EnumTipeAkun.Income)
            '2. pindahkan data COGS
            TotalCOGS = Me.PindahkanData(vTBLSumber, vTBLTarget, ClsMasterAkun.EnumTipeAkun.Cost_of_Goods_Sales)
            '3. masukkan data gros profit
            rowBaru = vTBLTarget.NewTrial_BalanceRow
            With rowBaru
                .Aktif = 1
                .Beginning_Balance = 0
                .Bold = 1
                .Credit = 0
                .Debet = TotalIncome - TotalCOGS
                If Me.chkUseTotal.Checked = True Then
                    .Akun = "Gross Profit"
                    .Description = " "
                Else
                    .Akun = " "
                    .Description = "Gross Profit"
                End If
                .Header = 0
                .Level_Perkiraan = 1
                .Primary_Akun = Me.GetPrimaryAkunMaks(vTBLSumber) + 1
                .Saldo = 0
                .Saldo_Normal = 0
                .Tipe_Perkiraan = 9 '--> tipe perkiraan baru, supaya bisa digrouping pada laporan
                .Total = 0
                .Urutan = 9
                .Account_Code = .Akun
            End With
            vTBLTarget.Rows.Add(rowBaru)
            '4. pindahkan data expense
            TotalExpense = Me.PindahkanData(vTBLSumber, vTBLTarget, ClsMasterAkun.EnumTipeAkun.Expense)
            '5. Pindahkan data other income
            TotalOtherIncome = Me.PindahkanData(vTBLSumber, vTBLTarget, ClsMasterAkun.EnumTipeAkun.Other_Income)
            '6. Pindahkan data other expense
            TotalOtherExpense = Me.PindahkanData(vTBLSumber, vTBLTarget, ClsMasterAkun.EnumTipeAkun.Other_Expense)
            '7. Masukkan data net profit before tax
            rowBaru = vTBLTarget.NewTrial_BalanceRow
            With rowBaru
                .Aktif = 1
                .Beginning_Balance = 0
                .Bold = 1
                .Credit = 0
                .Debet = TotalIncome - TotalCOGS - TotalExpense + TotalOtherIncome - TotalOtherExpense
                If Me.chkUseTotal.Checked = True Then
                    .Akun = "Nett Profit Before Tax"
                    .Description = " "
                Else
                    .Akun = " "
                    .Description = "Nett Profit Before Tax"
                End If
                .Header = 0
                .Level_Perkiraan = 1
                .Primary_Akun = Me.GetPrimaryAkunMaks(vTBLTarget) + 1
                .Saldo = 0
                .Saldo_Normal = 0
                .Tipe_Perkiraan = 10 '--> tipe perkiraan baru, supaya bisa digrouping pada laporan
                .Total = 0
                .Urutan = 10
                .Account_Code = .Akun
            End With
            vTBLTarget.Rows.Add(rowBaru)
            '8. Pindahkan data Tax
            TotalTax = Me.PindahkanData(vTBLSumber, vTBLTarget, ClsMasterAkun.EnumTipeAkun.Tax)
            '9. Masukkan data net profit after tax
            rowBaru = vTBLTarget.NewTrial_BalanceRow
            With rowBaru
                .Aktif = 1
                .Beginning_Balance = 0
                .Bold = 1
                .Credit = 0
                .Debet = TotalIncome - TotalCOGS - TotalExpense + TotalOtherIncome - TotalOtherExpense - TotalTax
                If Me.chkUseTotal.Checked = True Then
                    .Akun = "Nett Profit After Tax"
                    .Description = " "
                Else
                    .Akun = " "
                    .Description = "Nett Profit After Tax"
                End If
                .Header = 0
                .Level_Perkiraan = 1
                .Primary_Akun = Me.GetPrimaryAkunMaks(vTBLTarget) + 1
                .Saldo = 0
                .Saldo_Normal = 0
                .Tipe_Perkiraan = 11 '--> tipe perkiraan baru, supaya bisa digrouping pada laporan
                .Total = 0
                .Urutan = 11
                .Account_Code = .Akun
            End With
            vTBLTarget.Rows.Add(rowBaru)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function PindahkanData(ByVal vTBLSumber As DseLaporanTrialBalance.Trial_BalanceDataTable, _
    ByVal vTBLTarget As DseLaporanTrialBalance.Trial_BalanceDataTable, ByVal TipeAkun As ClsMasterAkun.EnumTipeAkun) As Double
        Dim rowSumber, rowTarget As DseLaporanTrialBalance.Trial_BalanceRow
        Dim SubTotal As Double = 0

        Try
            'start pindahkan data
            '1. pindahkan data income
            For Each rowSumber In vTBLSumber.Rows
                If rowSumber.Tipe_Perkiraan = TipeAkun Then
                    rowTarget = vTBLTarget.NewTrial_BalanceRow
                    With rowTarget
                        .Aktif = rowSumber.Aktif
                        .Akun = rowSumber.Akun
                        .Beginning_Balance = rowSumber.Beginning_Balance
                        .Credit = rowSumber.Credit
                        .Debet = rowSumber.Debet
                        .Description = rowSumber.Description
                        .Header = rowSumber.Header
                        .Level_Perkiraan = rowSumber.Level_Perkiraan
                        .Primary_Akun = rowSumber.Primary_Akun
                        .Saldo = rowSumber.Saldo
                        .Saldo_Normal = rowSumber.Saldo_Normal
                        .Tipe_Perkiraan = rowSumber.Tipe_Perkiraan
                        .Urutan = rowSumber.Tipe_Perkiraan
                        .Account_Code = rowSumber.Account_Code
                    End With

                    If rowSumber.Saldo_Normal = 0 Then
                        SubTotal += rowSumber.Debet
                    Else
                        SubTotal += rowSumber.Credit
                    End If

                    vTBLTarget.Rows.Add(rowTarget)
                End If
            Next

            Return SubTotal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

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
                .Account_Code = row.Account_Code.Trim
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

    Private Sub Cetak_P_L(ByVal Adjusted As Boolean)
        Dim TBL As DseLaporanTrialBalance.Trial_BalanceDataTable = Me.DS.Trial_Balance
        Dim row As DseLaporanTrialBalance.Trial_BalanceRow
        Dim rowCetak As DseLaporanTrialBalance.TblCetakRow
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim i As Double = 0
        Dim vDtpFrom_1 As Date = Now
        Dim vDtpTo_1 As Date = Now
        Dim splitAkun(1) As String
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



            'For x As Integer = 0 To DS.TblCetak.Rows.Count - 1
            '    rowCetak = DS.TblCetak.Rows(x)
            '    rowCetak.Akun = rowCetak.Akun.Trim
            '    If Me.chkShowAccountCode.Checked = True Then
            '        splitAkun = Split(rowCetak.Akun, "-")
            '        If splitAkun.Length <> 1 Then
            '            rowCetak.Akun_No = splitAkun(0).Trim
            '            rowCetak.Akun_Name = splitAkun(1).Trim
            '            If splitAkun.Length > 2 Then
            '                For Counter As Integer = 2 To splitAkun.Length - 1
            '                    rowCetak.Akun_Name &= "-" & splitAkun(Counter)
            '                Next
            '            End If
            '        Else
            '            rowCetak.Akun_Name = rowCetak.Akun.Trim
            '        End If

            '    End If
            'Next
            Me.report.SetDataSource(DS)
            Me.report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, Me.Def_Prec)
            Me.report.SetParameterValue(report.Parameter_Tanggal_Dari.ParameterFieldName, vDtpFrom_1)
            Me.report.SetParameterValue(report.Parameter_Tanggal_Sampai.ParameterFieldName, vDtpTo_1)
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

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub Cetak_P_L_Comparison_Period(ByVal Adjusted As Boolean)
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
                    .Header = row.Header
                    If row.Saldo_Normal = 0 Then
                        .Saldo_Period_I = row.Debet
                    Else
                        .Saldo_Period_I = row.Credit
                    End If
                    ''TR//Perubahan Account jadi rata kiri sama misahin antara No / Name Account
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

            ''TR//Perubahan Account jadi rata kiri sama misahin antara No / Name Account
            'For x As Integer = 0 To DS.TblCetak.Rows.Count - 1
            '    rowCetak = DS.TblCetak.Rows(x)
            '    rowCetak.Akun = rowCetak.Akun.Trim
            '    If Me.chkShowAccountCode.Checked = True Then
            '        splitAkun = Split(rowCetak.Akun, "-")
            '        If splitAkun.Length <> 1 Then
            '            rowCetak.Akun_No = splitAkun(0).Trim
            '            rowCetak.Akun_Name = splitAkun(1).Trim
            '            If splitAkun.Length > 2 Then
            '                For Counter As Integer = 2 To splitAkun.Length - 1
            '                    rowCetak.Akun_Name &= "-" & splitAkun(Counter)
            '                Next
            '            End If
            '        Else
            '            rowCetak.Akun_Name = rowCetak.Akun
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
                Me.report_comparison.SetParameterValue(report_comparison.Parameter_ShowNoAccount.ParameterFieldName, 1)
            Else
                Me.report_comparison.SetParameterValue(report_comparison.Parameter_ShowNoAccount.ParameterFieldName, 0)
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

    Private Sub Cetak_P_L_Comparison_Budget(ByVal Adjusted As Boolean)
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
        Dim splitAkun(1) As String

        Try
            'untuk cetak trial balance normal, yang ditampilkan adalah nilai: BB + SL + PR
            'maksudnya:
            '1. Isikan tabel dengan fungsi yang sudah disediakan
            Me.IsiTabel(Adjusted, Me.dtpFrom_I.Value, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpTo_I.Value), TBL)
            Me.ObjBudget.GetBudget_PL(Me.ObjCD.GetLasDateOfTheMonth(Me.dtpFrom_I.Value), Me.ObjCD.GetLasDateOfTheMonth(Me.dtpTo_I.Value), TBL2)
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
                    ''TR//Perubahan Account jadi rata kiri sama misahin antara No / Name Account
                    
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

            ''TR//Perubahan Account jadi rata kiri sama misahin antara No / Name Account
            'For x As Integer = 0 To DS.TblCetak.Rows.Count - 1
            '    rowCetak = DS.TblCetak.Rows(x)
            '    rowCetak.Akun = rowCetak.Akun.Trim
            '    If Me.chkShowAccountCode.Checked = True Then
            '        splitAkun = Split(rowCetak.Akun, "-")
            '        If splitAkun.Length <> 1 Then
            '            rowCetak.Akun_No = splitAkun(0).Trim
            '            rowCetak.Akun_Name = splitAkun(1).Trim
            '            If splitAkun.Length > 2 Then
            '                For Counter As Integer = 2 To splitAkun.Length - 1
            '                    rowCetak.Akun_Name &= "-" & splitAkun(Counter)
            '                Next
            '            End If
            '        Else
            '            rowCetak.Akun_Name = rowCetak.Akun
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
                Me.report_comparison.SetParameterValue(report_comparison.Parameter_ShowNoAccount.ParameterFieldName, 1)
            Else
                Me.report_comparison.SetParameterValue(report_comparison.Parameter_ShowNoAccount.ParameterFieldName, 0)
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
        Dim MinDate As Date = Today
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
            If Me.dtpTo_I.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpTo_I.Value.Date Then
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
            If Me.dtpTo_II.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpTo_II.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            MinDate = Me.Obj_Closing.GetConvertionDate
            If Me.dtpFrom_I.Value.Date < MinDate.Date Then
                MsgBox("Minimum data is : " + _
                Format(MinDate, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If Me.dtpFrom_II.Enabled = True Then
                If Me.dtpFrom_II.Value.Date < MinDate.Date Then
                    MsgBox("Minimum data is : " + _
                    Format(MinDate, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.cboFormatLaporan.SelectedIndex = 0 Then
                If Me.dtpFrom_I.Value.Date > Me.dtpTo_I.Value.Date Then
                    MsgBox("Incorrect period I", MsgBoxStyle.Information, "Information")
                    Me.crv1.ReportSource = Nothing
                    Me.crv1.Refresh()
                    Exit Sub
                End If
                Me.Cetak_P_L(True)
            ElseIf Me.cboFormatLaporan.SelectedIndex = 1 Then
                If Me.dtpFrom_I.Value.Date > Me.dtpTo_I.Value.Date Then
                    MsgBox("Incorrect period I", MsgBoxStyle.Information, "Information")
                    Me.crv1.ReportSource = Nothing
                    Me.crv1.Refresh()
                    Exit Sub
                End If
                If Me.dtpFrom_II.Value.Date > Me.dtpTo_II.Value.Date Then
                    MsgBox("Incorrect period II", MsgBoxStyle.Information, "Information")
                    Me.crv1.ReportSource = Nothing
                    Me.crv1.Refresh()
                    Exit Sub
                End If
                Me.Cetak_P_L_Comparison_Period(True)
            ElseIf Me.cboFormatLaporan.SelectedIndex = 2 Then
                If Me.dtpFrom_I.Value.Date > Me.dtpTo_I.Value.Date Then
                    MsgBox("Incorrect period I", MsgBoxStyle.Information, "Information")
                    Me.crv1.ReportSource = Nothing
                    Me.crv1.Refresh()
                    Exit Sub
                End If
                Me.Cetak_P_L_Comparison_Budget(True)
            ElseIf Me.cboFormatLaporan.SelectedIndex = 3 Then
                Me.Cetak_P_L_2Period_YTD(True)
            ElseIf Me.cboFormatLaporan.SelectedIndex = 4 Or Me.cboFormatLaporan.SelectedIndex = 5 Then
                Me.Cetak_P_L_6MonthsPeriod(True)
            ElseIf Me.cboFormatLaporan.SelectedIndex = 6 Or Me.cboFormatLaporan.SelectedIndex = 7 Then
                Me.Cetak_P_L_12MonthsPeriod(True)
            ElseIf Me.cboFormatLaporan.SelectedIndex = 9 Then
                If Me.dtpFrom_I.Value.Date > Me.dtpTo_I.Value.Date Then
                    MsgBox("Incorrect period I", MsgBoxStyle.Information, "Information")
                    Me.crv1.ReportSource = Nothing
                    Me.crv1.Refresh()
                    Exit Sub
                End If
                Me.Cetak_P_L_ChartMonthly(True)
            Else
                Me.Cetak_P_L_QuarterYearly(True)
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

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
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