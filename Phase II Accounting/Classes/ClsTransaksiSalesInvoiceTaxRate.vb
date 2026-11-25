Option Explicit On
Option Strict Off

Public Class ClsTransaksiSalesInvoiceTaxRate
    Public Function GetTaxRate(ByVal vTanggal As Date, ByVal vPrimaryCurrency As Double) As Decimal
        Dim DTA As New DseTransaksiSalesInvoice_TaxRateTableAdapters.tbltrans_salesinvoice_taxrate_itemTableAdapter

        Try
            If DTA.ScalarQuery_GetRate(vPrimaryCurrency, vTanggal).HasValue = True Then
                Return DTA.ScalarQuery_GetRate(vPrimaryCurrency, vTanggal).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
