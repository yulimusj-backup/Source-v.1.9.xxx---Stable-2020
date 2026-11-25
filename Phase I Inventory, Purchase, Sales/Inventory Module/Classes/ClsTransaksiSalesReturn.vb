Option Explicit On
Option Strict Off

Public Class ClsTransaksiSalesReturn
    Public Enum EnumTipeSR As Integer
        Ganti_Barang
        Revisi_Invoice
        Nota_Retur_Pajak
    End Enum

    Public Sub GetNoTransaksi_Remarks_NoInvoice(ByRef vNoTransaksi As String, ByRef vRemarks As String, ByVal vPrimary As Double, ByRef vDate As Date)
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturn1TableAdapter
        Dim TBL As New dseTransaksiSalesReturn.tbltrans_salesreturn1DataTable
        Dim row As dseTransaksiSalesReturn.tbltrans_salesreturn1Row

        Try
            DTA.Fill(TBL, vPrimary)

            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                If row.No_Invoice <> "" Then
                    If row.Remarks <> "" Then
                        vRemarks = "SR Number: " + row.No_Transaksi + " - " + row.Remarks
                    Else
                        vRemarks = "SR Number: " + row.No_Transaksi
                    End If
                    vNoTransaksi = row.No_Invoice
                Else
                    vRemarks = row.Remarks
                    vNoTransaksi = row.No_Transaksi
                End If
                vDate = row.Tanggal_Transaksi
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

    Public Function GetNoSalesReturn(ByVal vPrimaryStockInSalesReturn As Double) As String
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.DataTable1TableAdapter
        Dim TBL As New dseTransaksiSalesReturn.DataTable1DataTable
        Dim row As dseTransaksiSalesReturn.DataTable1Row

        Try
            If DTA.Fill(TBL, vPrimaryStockInSalesReturn) > 0 Then
                row = TBL.Rows(0)
                Return row.No_Transaksi
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function
End Class
