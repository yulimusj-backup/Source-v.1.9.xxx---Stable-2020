Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Public Class ClsTransaksiLoanPayment_Item

    Public Enum EnumNilaiUpdate As Integer
        NilaiDefault
        Update
    End Enum

    Public Function GetJmlLoanItemByIdentifier(ByVal vIdentifierLoan As String) As Double
        Dim DTA As New DseTransaksiLoanPaymentTableAdapters.tbltrans_loan_itemTableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetJmlByIdentifier(vIdentifierLoan)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetIdentifierLoan(ByVal vPrimaryTransaksi As Double) As String
        Dim DTA As New DseTransaksiLoanPaymentTableAdapters.tbltrans_loan_payment_temTableAdapter
        Dim Dummy As String
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetIdentifierLoan(vPrimaryTransaksi)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal vPrimary As Double, ByVal vDGVItem As DataGridView, _
    ByVal vDef_Prec As Double, ByVal i As FrmTransaksiLoan_Payment) As Boolean
        Dim rowDGV As DataGridViewRow
        Dim DTAItem As New DseTransaksiLoanPaymentTableAdapters.DataTable1TableAdapter
        Dim TBLItem As New DseTransaksiLoanPayment.DataTable1DataTable
        Dim RowItem As DseTransaksiLoanPayment.DataTable1Row
        Dim ObjCD As New Core_Module.ClsCekData
        Try
            'update connection string
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'isikan data item overhead cost
            DTAItem.FillBy_PrimaryTransaksi(TBLItem, vPrimary)
            vDGVItem.Rows.Clear()
            For Each RowItem In TBLItem.Rows
                vDGVItem.Rows.Add()
                rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                With rowDGV
                    .Cells(i.ColNo.Index).Value = ""
                    .Cells(i.ColLoanNo.Index).Value = RowItem.Loan_Number
                    .Cells(i.ColInstallementAt.Index).Value = RowItem.Installment_At
                    .Cells(i.ColBase.Index).Value = ObjCD.DigitGrouping(RowItem.Base, vDef_Prec)
                    .Cells(i.ColInterest.Index).Value = ObjCD.DigitGrouping(RowItem.Interest, vDef_Prec)
                    .Cells(i.ColTotalInstallment.Index).Value = ObjCD.DigitGrouping(RowItem.Base + RowItem.Interest, vDef_Prec)
                    .Cells(i.ColPayment.Index).Value = ObjCD.DigitGrouping(RowItem.Payment, vDef_Prec)
                    .Cells(i.ColRemarks.Index).Value = RowItem.Remarks
                    .Cells(i.ColPrimaryLoan.Index).Value = RowItem.Primary_Loan
                    .Cells(i.ColIdentifier.Index).Value = RowItem.Identifier_Loan
                    .Cells(i.ColPrimaryRev.Index).Value = RowItem.Primary_Receivable_Account
                End With
            Next

            'set number
            ObjCD.SetNumber(vDGVItem)

            'autosize dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            TBLItem.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function HapusItem(ByVal vPrimary As Double) As Boolean
        Dim DTAItem As New DseTransaksiLoanPaymentTableAdapters.tbltrans_loan_payment_temTableAdapter
        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.Delete(vPrimary)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
        End Try
    End Function

    Public Function SimpanItem(ByVal vPrimary As Double, _
    ByVal vDGVItem As DataGridView, ByVal i As FrmTransaksiLoan_Payment) As Boolean
        Dim DTAItem As New DseTransaksiLoanPaymentTableAdapters.tbltrans_loan_payment_temTableAdapter
        Dim rowDGV As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData
        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For Each rowDGV In vDGVItem.Rows
                With rowDGV
                    DTAItem.Insert(vPrimary, .Cells(i.ColPrimaryLoan.Index).Value, _
                    .Cells(i.ColIdentifier.Index).Value, .Cells(i.ColPayment.Index).Value, _
                    .Cells(i.ColRemarks.Index).Value)
                End With
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function ApakahLoanNumberBolehDiHapus(ByVal vPrimaryLoanMin As Double, ByVal i As FrmTransaksiLoan_Payment) As String
        Dim DTA As New DseTransaksiLoanPaymentTableAdapters.tbltrans_loan_itemTableAdapter
        Dim TBL As New DseTransaksiLoanPayment.tbltrans_loan_itemDataTable
        Dim Row As DseTransaksiLoanPayment.tbltrans_loan_itemRow
        Dim DTAPay As New DseTransaksiLoanPaymentTableAdapters.tbltrans_loan_payment_temTableAdapter
        Dim TBLPay As New DseTransaksiLoanPayment.tbltrans_loan_payment_temDataTable
        Dim RowPay As DseTransaksiLoanPayment.tbltrans_loan_payment_temRow
        Dim RowDGV As DataGridViewRow
        Dim Filter As String = ""
        Dim Retval As String = ""
        Dim MaxCicilan As Double = 0
        Dim ij As Integer = 0
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAPay.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_PrimaryLoanMinimumNPaid(TBL, vPrimaryLoanMin)
            DTAPay.FillBy_PrimaryTransaksi(TBLPay, i.TxtPrimaryLoanPayment.Text)
            'Primary dirinya sendiri tidak termasuk ke dalam pengecekan
            If i.TxtPrimaryLoanPayment.Text <> "-1" Then
                If TBLPay.Rows.Count > 0 Then
                    For Each RowPay In TBLPay.Rows
                        Filter = "[Identifier] = '" + RowPay.Identifier_Loan.ToString + "'"
                        TBL.DefaultView.RowFilter = Filter
                        If TBL.DefaultView.Count > 0 Then
                            Row = TBL.DefaultView(0).Row
                            TBL.Rows.Remove(Row)
                        End If
                    Next
                End If
            End If
            'Pengecekan maximum Loan yang telah dibayar
            For Each RowDGV In i.dgvItem.Rows
                MaxCicilan = 0
                TBL.DefaultView.RowFilter = ""
                TBL.DefaultView.Sort = ""
                Filter = "[Primary Transaksi] = " + RowDGV.Cells(i.ColPrimaryLoan.Index).Value.ToString + ""
                TBL.DefaultView.RowFilter = Filter
                TBL.DefaultView.Sort = "[Primary Transaksi], [Identifier]"
                If TBL.DefaultView.Count > 0 Then
                    Row = TBL.DefaultView(TBL.DefaultView.Count - 1).Row
                    MaxCicilan = Row.Installment_Period
                    If CDbl(RowDGV.Cells(i.ColInstallementAt.Index).Value) < MaxCicilan Then
                        ij += 1
                        If Retval = "" Then
                            Retval = ij.ToString + ". Loan number : " + RowDGV.Cells(i.ColLoanNo.Index).Value.ToString + ", Installment At : " + RowDGV.Cells(i.ColInstallementAt.Index).Value.ToString + "."
                        Else
                            Retval += vbCrLf + ij.ToString + ". Loan number : " + RowDGV.Cells(i.ColLoanNo.Index).Value.ToString + ", Installment At : " + RowDGV.Cells(i.ColInstallementAt.Index).Value.ToString + "."
                        End If
                    End If
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTAPay.Dispose()
            TBLPay.Dispose()
        End Try
    End Function

    Public Sub UpdatePaymentLoan_Item(ByVal vBaru As Boolean, ByVal vPaymentDate As Date, ByVal TBLLoan As DseTransaksiLoanPayment.DataTable2DataTable, _
    ByVal vFungsiPengunaan As EnumNilaiUpdate, ByVal i As FrmTransaksiLoan_Payment)
        Dim DTAPinjam As New DseTransaksiLoanPaymentTableAdapters.tbltrans_loan_itemTableAdapter
        Dim RowLoan As DseTransaksiLoanPayment.DataTable2Row
        Dim RowDGV As DataGridViewRow
        Dim RV As DataRowView
        Dim UpdateField As Integer = 1
        Dim PaymentDateDefault As Date = #1/1/1950#
        Dim PaymentValueDefault As Double = 0
        Dim Filter As String
        Try
            DTAPinjam.Connection.ConnectionString = Core_Module.GetConstring
            If vBaru = True Then
                If TBLLoan.Rows.Count > 0 Then
                    Filter = "[Update] = " + UpdateField.ToString + ""
                    TBLLoan.DefaultView.RowFilter = Filter
                    If TBLLoan.DefaultView.Count > 0 Then
                        For Each RV In TBLLoan.DefaultView
                            RowLoan = RV.Row
                            If vFungsiPengunaan = EnumNilaiUpdate.NilaiDefault Then
                                DTAPinjam.UpdateQuery_NilaiPayment(PaymentDateDefault, PaymentValueDefault, RowLoan.Identifier)
                            ElseIf vFungsiPengunaan = EnumNilaiUpdate.Update Then
                                If RowLoan.Payment_Amount = 0 Then
                                    DTAPinjam.UpdateQuery_NilaiPayment(PaymentDateDefault, PaymentValueDefault, RowLoan.Identifier)
                                Else
                                    DTAPinjam.UpdateQuery_NilaiPayment(vPaymentDate, RowLoan.Payment_Amount, RowLoan.Identifier)
                                End If
                            End If
                        Next
                    End If
                    TBLLoan.DefaultView.RowFilter = ""
                End If
            End If
            If vBaru = False Then
                If i.dgvItem.Rows.Count > 0 Then
                    For Each RowDGV In i.dgvItem.Rows
                        DTAPinjam.UpdateQuery_NilaiPayment(i.dtpDate.Value.Date, RowDGV.Cells(i.ColPayment.Index).Value, _
                        RowDGV.Cells(i.ColIdentifier.Index).Value)
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAPinjam.Dispose()
        End Try
    End Sub

    Public Function GetPrimaryReceivableAcc_Loan(ByVal vPrimaryLoan As Double) As Double
        Dim DTA As New DseTransaksiLoanPaymentTableAdapters.tbltrans_loan_paymentTableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPrimaryRecevable_Loan(vPrimaryLoan)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

End Class
