Option Explicit On
Option Strict Off

Public Class ClsTransaksiSalesInvoiceItem
    'elemen yang digunakan pada automatic invoice transfer
    Public Enum EnumTipeSalesBalance As Integer
        Automatic
        Manual
        Manual_Transfer '--> tidak digunakan lagi, hanya pada cmsi saja
        Automatic_SR
        Manual_SR
    End Enum

    Public Function GetJmlByDO(ByVal PrimaryDO As Double) As Integer
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_itemTableAdapter
        Try
            Return DTA.ScalarQuery_GetJmlByPrimaryDO(PrimaryDO).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetSalesInvoiceNo(ByVal vPrimaryDo As Double) As String
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_itemTableAdapter

        Try
            Return DTA.ScalarQuery_GetNoSalesInvoice(vPrimaryDo)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
