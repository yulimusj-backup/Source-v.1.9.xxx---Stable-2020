Option Explicit On
Option Strict Off
Imports AccModule
Imports System.IO
Imports System.Xml
Public Class FrmLaporanChequeAging
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjCurr As New Core_Module.clsMasterCurrency
    Dim ObjCheque As New ClsTransaksiJournal_Item_Cheque
    Dim TBL As New DseLaporanChequePerPeriod_Aging.tbltrans_jurnal_chequeDataTable
    Dim period_1 As Date = #1/1/1950#
    Dim period_2 As Date = #1/1/1950#
    Dim period_3 As Date = #1/1/1950#
    Dim period_lebih3month As Date = #1/1/1950#
    Dim ObjLicense As New Core_Module.ClsLicense

    Private Sub SesuaikanPeriod(ByVal DGV As DataGridView)
        Dim per1 As String
        Dim per2 As String
        Dim per3 As String
        Dim date2 As Date
        Dim date3 As Date
        Try
            per1 = Format(Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value.Date), "MMM yy")
            date2 = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value.Date + TimeSpan.FromDays(32))
            date3 = Me.ObjCD.GetFirstDateofTheMonth(date2 + TimeSpan.FromDays(32))
            per2 = Format(date2, "MMM yy")
            per3 = Format(date3, "MMM yy")
            DGV.Columns(Me.ColPeriod1.Index).HeaderText = per1
            DGV.Columns(Me.ColPeriod2.Index).HeaderText = per2
            DGV.Columns(Me.ColPeriod3.Index).HeaderText = per3
            DGV.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanChequeAging_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.dtpPeriod.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value)
            Me.cboType.SelectedIndex = 0
            Me.CancelButton = Me.btnClose
            Me.dtpPeriod.Enabled = True
            Me.SesuaikanPeriod(Me.dgvCheque)
            Me.dgvCheque.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.btnClose.Anchor = AnchorStyles.Right And AnchorStyles.Bottom
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'fungsi ini menyesuaikan header period 1 - 4
    Private Sub SetupPeriod(ByVal vInOutGoing As ClsTransaksiJournal.EnumTipeInOut, _
    ByVal vTBL As DseLaporanChequePerPeriod_Aging.tbltrans_jurnal_chequeDataTable, _
    Optional ByVal vPrint As Boolean = False)
        Dim DTA As New DseLaporanChequePerPeriod_AgingTableAdapters.tbltrans_jurnal_chequeTableAdapter
        Dim Row As DseLaporanChequePerPeriod_Aging.tbltrans_jurnal_chequeRow
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If vPrint = False Then
                vTBL.Rows.Clear()
            End If
            Me.period_1 = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value.Date)
            Me.period_2 = Me.ObjCD.GetFirstDateofTheMonth(period_1.Date + TimeSpan.FromDays(32))
            Me.period_3 = Me.ObjCD.GetFirstDateofTheMonth(period_2.Date + TimeSpan.FromDays(32))
            Me.period_lebih3month = Me.ObjCD.GetFirstDateofTheMonth(period_3 + TimeSpan.FromDays(32))
            DTA.FillBy_TypeInOut(vTBL, vInOutGoing)
            For Each Row In vTBL.Rows
                If Row.Due_Date.Date > Me.period_lebih3month.Date Then
                    Row.Period_4 = Row.Amount
                Else
                    If Row.Due_Date.Date > Me.period_3.Date Then
                        Row.Period_3 = Row.Amount
                    Else
                        If Row.Due_Date.Date > Me.period_2.Date Then
                            Row.Period_2 = Row.Amount
                        Else
                            If Row.Due_Date.Date > Me.period_1.Date Then
                                Row.Period_1 = Row.Amount
                            Else
                                Row.OverDue = Row.Amount
                            End If
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Private Sub LoadData_DGV(ByVal vInOutGoing As ClsTransaksiJournal.EnumTipeInOut, _
     ByVal vTBLCheque As DseLaporanChequePerPeriod_Aging.tbltrans_jurnal_chequeDataTable)
        Dim Row As DseLaporanChequePerPeriod_Aging.tbltrans_jurnal_chequeRow
        Dim rowDGV As DataGridViewRow
        Dim ij As TimeSpan
        Dim dec As Double = 0
        Try
            Me.dgvCheque.Rows.Clear()
            Me.dgvCheque.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvCheque.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            If Me.cboType.SelectedIndex = 0 Then
                ' In Coming 
                Me.SetupPeriod(ClsTransaksiJournal.EnumTipeInOut._In, vTBLCheque)
            Else
                ' Out Going 
                Me.SetupPeriod(ClsTransaksiJournal.EnumTipeInOut.out, vTBLCheque)
            End If
            For Each Row In vTBLCheque.Rows
                Me.dgvCheque.Rows.Add()
                rowDGV = Me.dgvCheque.Rows(Me.dgvCheque.Rows.Count - 1)
                rowDGV.Cells(Me.ColChequeNo.Index).Value = Row.Cheque_No.ToString
                rowDGV.Cells(Me.ColBank.Index).Value = Row.Bank.ToString
                rowDGV.Cells(Me.ColDate.Index).Value = Format(Row.Cheque_date.Date, "dd MMM yyyy")
                rowDGV.Cells(Me.ColDueDate.Index).Value = Format(Row.Due_Date.Date, "dd MMM yyyy")
                ij = Row.Due_Date.Date - Row.Cheque_date.Date
                rowDGV.Cells(Me.ColPaymentTerms.Index).Value = ij.Days
                rowDGV.Cells(Me.ColCurrency.Index).Value = Row.Currency_Code
                dec = Me.ObjCurr.GetDefaultPrecission(Row.Primary_Currency)
                rowDGV.Cells(Me.ColOverDue.Index).Value = Me.ObjCD.DigitGrouping(Row.OverDue, dec)
                rowDGV.Cells(Me.ColPeriod1.Index).Value = Me.ObjCD.DigitGrouping(Row.Period_1, dec)
                rowDGV.Cells(Me.ColPeriod2.Index).Value = Me.ObjCD.DigitGrouping(Row.Period_2, dec)
                rowDGV.Cells(Me.ColPeriod3.Index).Value = Me.ObjCD.DigitGrouping(Row.Period_3, dec)
                rowDGV.Cells(Me.ColPeriod4.Index).Value = Me.ObjCD.DigitGrouping(Row.Period_4, dec)
                rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value = Row.Primary_Currency
                rowDGV.Cells(Me.ColIdentifier.Index).Value = Row.Identifier
            Next
            Me.dgvCheque.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvCheque.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CetakLaporanChequeAging(ByVal vInOutGoing As ClsTransaksiJournal.EnumTipeInOut)
        Dim FormatDay As Integer = 1
        Dim FormatMonth As Integer = 2
        Dim FormatYears As Integer = 1
        'Dim FormatDecimal As Integer = 2
        Dim ij As New FrmLaporanChequeAging_Print
        Dim crv As CrystalDecisions.Windows.Forms.CrystalReportViewer = ij.crv1
        Dim Report As New CrpLaporanCheque_Aging
        Dim DS As New DseLaporanChequePerPeriod_Aging
        Dim RowLaporan As DseLaporanChequePerPeriod_Aging.tbltrans_jurnal_chequeRow
        Dim Judul As String
        Try
            If Me.cboType.SelectedIndex = 0 Then
                Me.SetupPeriod(ClsTransaksiJournal.EnumTipeInOut._In, DS.tbltrans_jurnal_cheque, True)
                Judul = "IN COMING CHEQUE AGING REPORT"
            Else
                Me.SetupPeriod(ClsTransaksiJournal.EnumTipeInOut.out, DS.tbltrans_jurnal_cheque, True)
                Judul = "OUT GOING CHEQUE AGING REPORT"
            End If
            For Each RowLaporan In DS.tbltrans_jurnal_cheque.Rows
                RowLaporan.Format_Decimal = Me.ObjCurr.GetDefaultPrecission(RowLaporan.Primary_Currency)
            Next

            Report.SetDataSource(DS)
            Report.SetParameterValue(Report.Parameter_Period_1.ParameterFieldName, Format(Me.period_1, "MMM yy"))
            Report.SetParameterValue(Report.Parameter_Period_2.ParameterFieldName, Format(Me.period_2, "MMM yy"))
            Report.SetParameterValue(Report.Parameter_Period_3.ParameterFieldName, Format(Me.period_3, "MMM yy"))
            Report.SetParameterValue(Report.Parameter_Tanggal.ParameterFieldName, Format(Me.dtpPeriod.Value.Date, "MMMM yyyy"))
            Report.SetParameterValue(Report.Parameter_Judul.ParameterFieldName, Judul)
            Report.SetParameterValue(Report.Parameter_Format_Day.ParameterFieldName, FormatDay)
            Report.SetParameterValue(Report.Parameter_Format_Month.ParameterFieldName, FormatMonth)
            Report.SetParameterValue(Report.Parameter_Format_Year.ParameterFieldName, FormatYears)
            crv.ReportSource = Report
            crv.Show()

            ij.Hide()
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            ij.Dispose()
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Dim EndLicenseDate As Date = #1/1/1950#
        Try
            EndLicenseDate = Me.ObjLicense.GetExpiryDate(My.Settings.b)
            If Me.dtpPeriod.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpPeriod.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            If Me.cboType.SelectedIndex = 0 Then
                Me.LoadData_DGV(ClsTransaksiJournal.EnumTipeInOut._In, Me.TBL)
            Else
                Me.LoadData_DGV(ClsTransaksiJournal.EnumTipeInOut.out, Me.TBL)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCheque_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvCheque.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.ColPay.Index Or e.ColumnIndex <> Me.ColViewJurnal.Index Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCheque_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvCheque.CellClick
        Dim Hasil As MsgBoxResult
        Dim objjurnal As New ClsTransaksiJournal
        Dim PrimaryJurnal As Double
        Dim TipeJurnal1 As ClsTransaksiJournal.EnumTipeJournal
        Dim TipeJurnal2 As ClsTransaksiJournal.EnumTipeJurnal2
        Try
            If e.ColumnIndex = Me.ColPay.Index Then
                Hasil = MsgBox("Do you want to pay this item ?", MsgBoxStyle.YesNo, "Confirm")
                If Hasil = MsgBoxResult.Yes Then
                    Me.ObjCheque.UpdatePaid_LapChequeAging(ClsTransaksiJournal_Item_Cheque.EnumPaid.Bayar, _
                    Me.dgvCheque.Item(Me.ColIdentifier.Index, e.RowIndex).Value)
                    If Me.cboType.SelectedIndex = 0 Then
                        Me.LoadData_DGV(ClsTransaksiJournal.EnumTipeInOut._In, Me.TBL)
                    Else
                        Me.LoadData_DGV(ClsTransaksiJournal.EnumTipeInOut.out, Me.TBL)
                    End If
                End If
            End If
            If e.ColumnIndex = Me.ColViewJurnal.Index Then
                'MsgBox("Belum", MsgBoxStyle.Information, "Information")
                PrimaryJurnal = Me.GetPrimaryByIdentifier(Me.dgvCheque.Item(Me.ColIdentifier.Index, e.RowIndex).Value)
                objjurnal.GetTipeJournal(PrimaryJurnal)
                TipeJurnal1 = objjurnal.GetTipeJournal(PrimaryJurnal)
                If TipeJurnal1 = ClsTransaksiJournal.EnumTipeJournal.GeneralJournal Then
                    TipeJurnal2 = objjurnal.GetTipeJournal2(PrimaryJurnal)
                    If TipeJurnal2 = ClsTransaksiJournal.EnumTipeJurnal2.Cash_n_Bank Then

                        Me.ShowCashBank(PrimaryJurnal)
                    End If
                End If
                If TipeJurnal1 = ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal Then
                    Me.ShowSalesReceivedJurnal(PrimaryJurnal)
                End If
                If TipeJurnal1 = ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Then
                    Me.ShowPurchasePaymentJurnal(PrimaryJurnal)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objjurnal = Nothing
        End Try
    End Sub

    Private Function GetPrimaryByIdentifier(ByVal vIdentifier As String) As Double
        Dim DTA As New DseLaporanChequePerPeriodTableAdapters.tbltrans_jurnal_chequeTableAdapter
        Try
            If DTA.ScalarQuery_GetPrimaryByIdentifer(vIdentifier).HasValue Then
                Return DTA.ScalarQuery_GetPrimaryByIdentifer(vIdentifier).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Sub btnPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Try
            If Me.dgvCheque.Rows.Count > 0 Then
                If Me.cboType.SelectedIndex = 0 Then
                    Me.CetakLaporanChequeAging(ClsTransaksiJournal.EnumTipeInOut._In)
                Else
                    Me.CetakLaporanChequeAging(ClsTransaksiJournal.EnumTipeInOut.out)
                End If
            Else
                MsgBox("Please filter data before print.", MsgBoxStyle.Information, "Information")
                Me.btnFilter.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpPeriod_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpPeriod.Leave
        Try
            Me.SesuaikanPeriod(Me.dgvCheque)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboType.SelectedIndexChanged
        Try
            Me.dgvCheque.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub ShowCashBank(ByVal vPrimaryJurnal As Double)
        Try
            Dim ij As AccModule.FrmTransaksiJournal_CashBank
            Dim j As Windows.Forms.Form

            For Each j In Me.ParentForm.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiJournal_CashBank)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    ij.ShowJurnal(vPrimaryJurnal)
                    Exit Sub
                End If
            Next

            Dim i As New AccModule.FrmTransaksiJournal_CashBank
            ' Increment the caption counter.
            i.UserCOde = Me.UserCode.ToString
            ' Set the caption to be unique.
            i.BackcolorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me.ParentForm
            i.Show()
            i.ShowJurnal(vPrimaryJurnal)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub ShowSalesReceivedJurnal(ByVal vPrimaryJurnal As Double)
        Try
            Dim ij As AccModule.FrmTransaksiJournal_Sales
            Dim j As Windows.Forms.Form

            For Each j In Me.ParentForm.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiJournal_Sales)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    ij.ShowJurnal(vPrimaryJurnal)
                    Exit Sub
                End If
            Next

            Dim i As New AccModule.FrmTransaksiJournal_Sales
            ' Increment the caption counter.
            i.UserCode = Me.UserCode.ToString
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me.ParentForm
            i.Show()
            i.ShowJurnal(vPrimaryJurnal)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub ShowPurchasePaymentJurnal(ByVal vPrimaryJurnal As Double)
        Try
            Dim ij As AccModule.FrmTransaksiJournal_Purchase
            Dim j As Windows.Forms.Form

            For Each j In Me.ParentForm.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiJournal_Purchase)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    ij.ShowJurnal(vPrimaryJurnal)
                    Exit Sub
                End If
            Next

            Dim i As New AccModule.FrmTransaksiJournal_Purchase
            ' Increment the caption counter.
            i.UserCode = Me.UserCode.ToString
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me.ParentForm
            i.Show()
            i.ShowJurnal(vPrimaryJurnal)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCheque_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvCheque.CellContentClick

    End Sub
End Class