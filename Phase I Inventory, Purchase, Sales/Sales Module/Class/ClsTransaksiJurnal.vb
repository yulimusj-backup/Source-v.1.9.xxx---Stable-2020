Option Explicit On
Option Strict Off

Public Class ClsTransaksiJurnal
    Public Enum EnumTipeSalesBalance As Integer
        Automatic
        Manual
        Manual_Transfer_Invoice
        Automatic_SR
        Manual_SR
    End Enum
    Public Enum EnumTipePayment As Integer
        SalesInvoice_SalesInvoiceDP
        SalesReturn
        AOP_AOPDP
        Purchase_Return
    End Enum

    Public Function ApakahSudahAda(ByVal vJournalNo As String, ByVal vDTA As DseTransaksiSalesInvoice_TransfJurnalTableAdapters.tbltrans_jurnalTableAdapter) As Boolean
        Dim Jml As Nullable(Of Long)

        Try
            Jml = vDTA.ScalarQuery_GetJml(vJournalNo)
            If Jml.HasValue = True Then
                If Jml.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function Get_SalesInvoiceJurnal(ByVal vPrimarySalesInvoice As Double) As String
        Dim DTA As New DseTransaksiJurnalTableAdapters.tbltrans_jurnal_item_paymentTableAdapter
        Dim TBL As New DseTransaksiJurnal.tbltrans_jurnal_item_paymentDataTable
        Dim row As DseTransaksiJurnal.tbltrans_jurnal_item_paymentRow
        Dim retval As String = ""
        Dim TipePayment As EnumTipePayment = EnumTipePayment.SalesInvoice_SalesInvoiceDP
        Dim TipeSalesBalance As EnumTipeSalesBalance = EnumTipeSalesBalance.Automatic

        Try
            DTA.FillBy(TBL, TipePayment, vPrimarySalesInvoice, TipeSalesBalance)
            For Each row In TBL.Rows
                If retval = "" Then
                    retval = row.Ref_Number
                Else
                    retval = retval + ", " + row.Ref_Number
                End If
            Next

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Get_SalesInvoiceJurnal_SR(ByVal vPrimarySR As Double) As String
        Dim DTA As New DseTransaksiJurnalTableAdapters.tbltrans_jurnal_item_paymentTableAdapter
        Dim TBL As New DseTransaksiJurnal.tbltrans_jurnal_item_paymentDataTable
        Dim row As DseTransaksiJurnal.tbltrans_jurnal_item_paymentRow
        Dim retval As String = ""
        Dim TipePayment As EnumTipePayment = EnumTipePayment.SalesInvoice_SalesInvoiceDP
        Dim TipeSalesBalance As EnumTipeSalesBalance = EnumTipeSalesBalance.Automatic_SR

        Try
            DTA.FillBy(TBL, TipePayment, vPrimarySR, TipeSalesBalance)
            For Each row In TBL.Rows
                If retval = "" Then
                    retval = row.Ref_Number
                Else
                    retval = retval + ", " + row.Ref_Number
                End If
            Next

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Get_SalesReturnJurnal(ByVal vPrimarySalesReturn As Double) As String
        Dim DTA As New DseTransaksiJurnalTableAdapters.tbltrans_jurnal_item_paymentTableAdapter
        Dim TBL As New DseTransaksiJurnal.tbltrans_jurnal_item_paymentDataTable
        Dim row As DseTransaksiJurnal.tbltrans_jurnal_item_paymentRow
        Dim retval As String = ""
        Dim TipePayment As EnumTipePayment = EnumTipePayment.SalesReturn
        Dim TipeSalesBalance As EnumTipeSalesBalance = EnumTipeSalesBalance.Automatic

        Try
            DTA.FillBy(TBL, TipePayment, vPrimarySalesReturn, TipeSalesBalance)
            For Each row In TBL.Rows
                If retval = "" Then
                    retval = row.Ref_Number
                Else
                    retval = retval + ", " + row.Ref_Number
                End If
            Next

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
