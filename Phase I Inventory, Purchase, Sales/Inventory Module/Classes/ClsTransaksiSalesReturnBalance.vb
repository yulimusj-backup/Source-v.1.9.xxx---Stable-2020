Option Explicit On
Option Strict Off

Public Class ClsTransaksiSalesReturnBalance
    Public Sub GetNoTransaksi_Remarks_Tanggal(ByRef vNoTransaksi As String, ByRef vRemarks As String, ByVal vDate As Date, ByVal vPrimary As Double)
        Dim DTA As New DseTransaksiSalesReturnBalanceTableAdapters.tbltrans_historicalsalesreturn_manualTableAdapter
        Dim TBL As New DseTransaksiSalesReturnBalance.tbltrans_historicalsalesreturn_manualDataTable
        Dim row As DseTransaksiSalesReturnBalance.tbltrans_historicalsalesreturn_manualRow

        Try
            DTA.FillBy_Primary(TBL, vPrimary)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                vRemarks = row.Remarks
                vNoTransaksi = row.AOSR_Number
                vDate = row.AOSR_Date
            Else
                vRemarks = ""
                vNoTransaksi = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub
End Class
