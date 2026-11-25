Option Explicit On
Option Strict Off

Public Class FrmLaporanChequePerPeriod
    Public UserCode As String = ""
    Dim frmLap As New FrmLaporanChequePerPeriod_Print
    Dim ObjCd As New Core_Module.ClsCekData
    Dim ObjCurr As New Core_Module.clsMasterCurrency
    Dim ObjLicense As New Core_Module.ClsLicense

    Private Sub chkDate_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkDate.CheckedChanged
        Try
            If Me.chkDate.Checked = True Then
                Me.dtpFrom.Enabled = True
                Me.dtpTo.Enabled = True
            Else
                Me.dtpFrom.Enabled = False
                Me.dtpTo.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        Try
            If Me.CheckBox2.Checked = True Then
                Me.DateTimePicker1.Enabled = True
                Me.DateTimePicker2.Enabled = True
            Else
                Me.DateTimePicker1.Enabled = False
                Me.DateTimePicker2.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiDgv()
        Dim DTA As New DseLaporanChequePerPeriodTableAdapters.tbltrans_jurnal_chequeTableAdapter
        Dim TBL As New DseLaporanChequePerPeriod.tbltrans_jurnal_chequeDataTable
        Dim row As DseLaporanChequePerPeriod.tbltrans_jurnal_chequeRow
        Dim paid As Boolean = True
        Dim objcurr As New Core_Module.clsMasterCurrency
        Dim objcd As New Core_Module.ClsCekData
        Dim status As Integer = 0
        Dim period As Integer = 0
        Dim tipe As Integer = 0
        Try
            Me.dgvCheque.Rows.Clear()
            'cari type
            If Me.cboType.SelectedItem = "In Coming" Then
                tipe = 0
            ElseIf Me.cboType.SelectedItem = "Out Going" Then
                tipe = 1
            Else
                Exit Sub
            End If

            'cari status paid
            If Me.chkPaid.Checked = True And Me.CheckBox1.Checked = False Then
                'paid
                status = 0
            ElseIf Me.chkPaid.Checked = False And Me.CheckBox1.Checked = True Then
                'unpaid
                status = 1
            Else
                'paid dan unpaid
                status = 2
            End If

            'cari tipe period
            If Me.chkDate.Checked = True And Me.CheckBox2.Checked = False Then
                'period berdasarkan cheqeu date
                period = 1
            ElseIf Me.chkDate.Checked = False And Me.CheckBox2.Checked = True Then
                'period berdasarkan due date
                period = 2
            ElseIf Me.chkDate.Checked = True And Me.CheckBox2.Checked = True Then
                'semua period aktif
                period = 3
            Else
                'semua period tidak aktif
                period = 4
            End If

            If status = 0 Or status = 1 Then
                If period = 1 Then
                    'berdasarkan cheque date period
                    DTA.FillBy_TypePaidDate(TBL, tipe, status, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date)
                ElseIf period = 2 Then
                    'berdasarkan due date
                    DTA.FillBy_TypePaidDueDate(TBL, tipe, status, Me.DateTimePicker2.Value.Date, Me.DateTimePicker1.Value.Date)
                ElseIf period = 3 Then
                    'berdasarkan semua date
                    DTA.FillBy_TypePaidAllDate(TBL, tipe, status, Me.dtpFrom.Value.Date, Me.DateTimePicker2.Value.Date, _
                     Me.dtpTo.Value.Date, Me.DateTimePicker1.Value.Date)
                ElseIf period = 4 Then
                    DTA.FillBy_TypeAndPaid(TBL, tipe, status)
                End If
            ElseIf status = 2 Then
                'keluarkan semua
                If period = 1 Then
                    'berdasarkan cheque date period
                    DTA.FillBy_TypeDate(TBL, tipe, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date)
                ElseIf period = 2 Then
                    'berdasarkan due date
                    DTA.FillBy_TypeDueDate(TBL, tipe, Me.DateTimePicker2.Value.Date, Me.DateTimePicker1.Value.Date)
                ElseIf period = 3 Then
                    'berdasarkan semua date
                    DTA.FillBy_TypeAllDate(TBL, tipe, Me.dtpFrom.Value.Date, Me.DateTimePicker2.Value.Date, _
                    Me.dtpTo.Value.Date, Me.DateTimePicker1.Value.Date)
                ElseIf period = 4 Then
                    DTA.FillBy_Type(TBL, tipe)
                End If
            End If

            'pindahkan ke DGV
            For Each row In TBL.Rows
                With row
                    If .Paid = 0 Then
                        paid = True
                    Else
                        paid = False
                    End If
                    Me.dgvCheque.Rows.Add(.Cheque_No, .Bank, Format(.Cheque_date, "dd MMM yyyy"), Format(.Due_Date, "dd MMM yyyy"), objcurr.GetCurrencyCode(.Primary_Currency), _
                    objcd.DigitGrouping(.Amount.ToString, objcurr.GetDefaultPrecission(.Primary_Currency)), .Remarks, paid, .Identifier, .Primary_Currency)
                End With
            Next
            If Me.cboType.SelectedIndex = 0 Then
                Me.PrintInComing(frmLap.crv1, Me.chkDate.Checked, Me.CheckBox2.Checked)
            Else
                Me.PrintOutGoing(frmLap.crv1, Me.chkDate.Checked, Me.CheckBox2.Checked)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            objcd = Nothing
            objcurr = Nothing
        End Try
    End Sub

    Private Sub PrintInComing(ByVal vCRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vDatePeriod As Boolean, ByVal vDuePeriod As Boolean)
        Dim rpt As New CrpLaporanCheque_INComing
        Dim DS As New DseLaporanChequePerPeriod
        Dim TBLLap As DseLaporanChequePerPeriod.tblcetak_periodDataTable = DS.tblcetak_period
        Dim rowLap As DseLaporanChequePerPeriod.tblcetak_periodRow
        Dim rowDGV As DataGridViewRow
        Try
            For Each rowDGV In Me.dgvCheque.Rows
                rowLap = TBLLap.NewRow
                rowLap.Due_Date = rowDGV.Cells(Me.ColDueDate.Index).Value
                If rowDGV.Cells(Me.ColSPaid.Index).Value = True Then
                    rowLap.Status = "Paid"
                Else
                    rowLap.Status = "Unpaid"
                End If
                rowLap.Cheque_Date = rowDGV.Cells(Me.ColDate.Index).Value
                rowLap.Bank = rowDGV.Cells(Me.ColBank.Index).Value
                rowLap.No_Cheque = rowDGV.Cells(Me.ColChequeNo.Index).Value
                rowLap.Curr = Me.ObjCurr.GetCurrencyCode(rowDGV.Cells(Me.ColPrimaryCurr.Index).Value)
                rowLap.Amount = Me.ObjCd.DigitGrouping(rowDGV.Cells(Me.ColAmount.Index).Value, _
                Me.ObjCurr.GetDefaultPrecission(rowDGV.Cells(Me.ColPrimaryCurr.Index).Value))
                rowLap.Remarks = rowDGV.Cells(Me.ColRemarks.Index).Value
                TBLLap.Addtblcetak_periodRow(rowLap)
            Next

            rpt.SetDataSource(DS)
            rpt.SetParameterValue(rpt.Parameter_Date_Period.ParameterFieldName, vDatePeriod)
            rpt.SetParameterValue(rpt.Parameter_Due_Period.ParameterFieldName, vDuePeriod)
            rpt.SetParameterValue(rpt.Parameter_Date_From.ParameterFieldName, Format(Me.dtpFrom.Value.Date, "MMMM yyyy "))
            rpt.SetParameterValue(rpt.Parameter_Date_To.ParameterFieldName, Format(Me.dtpTo.Value.Date, "MMMM yyyy "))
            rpt.SetParameterValue(rpt.Parameter_Due_Date_From.ParameterFieldName, Format(Me.DateTimePicker2.Value.Date, "MMMM yyyy "))
            rpt.SetParameterValue(rpt.Parameter_Due_Date_To.ParameterFieldName, Format(Me.DateTimePicker1.Value.Date, "MMMM yyyy "))

            vCRV.ReportSource = rpt
            vCRV.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            TBLLap.Dispose()
        End Try
    End Sub

    Private Sub PrintOutGoing(ByVal vCRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vDatePeriod As Boolean, ByVal vDuePeriod As Boolean)
        Dim rpt As New CrpLaporanCheque_OUTGoing
        Dim DS As New DseLaporanChequePerPeriod
        Dim TBLLap As DseLaporanChequePerPeriod.tblcetak_periodDataTable = DS.tblcetak_period
        Dim rowLap As DseLaporanChequePerPeriod.tblcetak_periodRow
        Dim rowDGV As DataGridViewRow
        Try
            For Each rowDGV In Me.dgvCheque.Rows
                rowLap = TBLLap.NewRow
                rowLap.Due_Date = rowDGV.Cells(Me.ColDueDate.Index).Value
                If rowDGV.Cells(Me.ColSPaid.Index).Value = True Then
                    rowLap.Status = "Paid"
                Else
                    rowLap.Status = "Unpaid"
                End If
                rowLap.Cheque_Date = rowDGV.Cells(Me.ColDate.Index).Value
                rowLap.Bank = rowDGV.Cells(Me.ColBank.Index).Value
                rowLap.No_Cheque = rowDGV.Cells(Me.ColChequeNo.Index).Value
                rowLap.Curr = Me.ObjCurr.GetCurrencyCode(rowDGV.Cells(Me.ColPrimaryCurr.Index).Value)
                rowLap.Amount = Me.ObjCd.DigitGrouping(rowDGV.Cells(Me.ColAmount.Index).Value, _
                Me.ObjCurr.GetDefaultPrecission(rowDGV.Cells(Me.ColPrimaryCurr.Index).Value))
                rowLap.Remarks = rowDGV.Cells(Me.ColRemarks.Index).Value
                TBLLap.Addtblcetak_periodRow(rowLap)
            Next

            rpt.SetDataSource(DS)
            rpt.SetParameterValue(rpt.Parameter_Date_Period.ParameterFieldName, vDatePeriod)
            rpt.SetParameterValue(rpt.Parameter_Due_Period.ParameterFieldName, vDuePeriod)
            rpt.SetParameterValue(rpt.Parameter_Date_From.ParameterFieldName, Format(Me.dtpFrom.Value.Date, "MMMM yyyy "))
            rpt.SetParameterValue(rpt.Parameter_Date_To.ParameterFieldName, Format(Me.dtpTo.Value.Date, "MMMM yyyy "))
            rpt.SetParameterValue(rpt.Parameter_Due_Date_From.ParameterFieldName, Format(Me.DateTimePicker2.Value.Date, "MMMM yyyy "))
            rpt.SetParameterValue(rpt.Parameter_Due_Date_To.ParameterFieldName, Format(Me.DateTimePicker1.Value.Date, "MMMM yyyy "))

            vCRV.ReportSource = rpt
            vCRV.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            TBLLap.Dispose()
        End Try
    End Sub

    Private Sub btnFilter_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Dim EndLicenseDate As Date = #1/1/1950#
        Try
            EndLicenseDate = Me.ObjLicense.GetExpiryDate(My.Settings.b)
            If Me.dtpFrom.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpFrom.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.dtpTo.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpTo.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            Me.IsiDgv()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim EndLicenseDate As Date = #1/1/1950#
        Try
            EndLicenseDate = Me.ObjLicense.GetExpiryDate(My.Settings.b)
            If Me.dtpFrom.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpFrom.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.dtpTo.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpTo.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            Me.frmLap.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCheque_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvCheque.CellBeginEdit
        Try
            'yang bisa diedit adalah:
            '1. Modify Payment

            If e.ColumnIndex <> Me.ColPaid.Index Then
                e.Cancel = True
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCheque_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvCheque.CellClick
        Dim DTA As New DseLaporanChequePerPeriodTableAdapters.tbltrans_jurnal_chequeTableAdapter
        Dim TBl As New DseLaporanChequePerPeriod.tblcetak_periodDataTable
        Dim objjurnal As New ClsTransaksiJournal
        Dim PrimaryJurnal As Double
        Dim TipeJurnal1 As ClsTransaksiJournal.EnumTipeJournal
        Dim TipeJurnal2 As ClsTransaksiJournal.EnumTipeJurnal2
        Try
            If e.ColumnIndex = Me.ColPaid.Index Then
                If Me.dgvCheque.Item(Me.ColSPaid.Index, e.RowIndex).Value = True Then
                    Me.dgvCheque.Item(Me.ColSPaid.Index, e.RowIndex).Value = False
                ElseIf Me.dgvCheque.Item(Me.ColSPaid.Index, e.RowIndex).Value = False Then
                    Me.dgvCheque.Item(Me.ColSPaid.Index, e.RowIndex).Value = True
                End If
                'update berdasarkan identifier
                If Me.dgvCheque.Item(Me.ColSPaid.Index, e.RowIndex).Value = True Then
                    DTA.UpdateQuery_Paid(0, Me.dgvCheque.Item(Me.ColIdentifier.Index, e.RowIndex).Value)
                ElseIf Me.dgvCheque.Item(Me.ColSPaid.Index, e.RowIndex).Value = False Then
                    DTA.UpdateQuery_Paid(1, Me.dgvCheque.Item(Me.ColIdentifier.Index, e.RowIndex).Value)
                End If
            End If

            If e.ColumnIndex = Me.ColViewJurnal.Index Then
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
            DTA.Dispose()
            TBl.Dispose()
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

    Private Sub FrmLaporanChequePerPeriod_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.btnClose.Anchor = AnchorStyles.Right And AnchorStyles.Bottom
            Me.btnClose.Location = New Point(Me.btnClose.Location.X + 125, Me.btnClose.Location.Y)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class