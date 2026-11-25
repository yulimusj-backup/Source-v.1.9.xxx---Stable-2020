Option Explicit On
Option Strict Off

Public Class clsTransaksiStockInOut1


    'elemen yang digunakan pada module asset
    Public Enum EnumInOutAdjust As Integer
        NoTipe
        OutAdjustMaterialUsageAssetAcquisition
        OutAdjustMaterialUsageAssetRevaluation
        InOutCustom
    End Enum


    Public Enum EnumTipeJournal As Integer
        GeneralJournal
        AdjustmenJournal
        ClosingJournal
        SalesJournal
        SalesReturnJournal
        SalesDownPayment
        PurchaseJournal
        PurchaseReturnJournal
        PurchaseDownPaymentJournal
        SalesReceivementJournal
        SalesReturnPaymentJournal
        PurchasePaymentJournal
        PurchaseReturnReceivement
    End Enum

    Public Enum EnumTipeStockInOut As Integer
        StockIn_Purchase
        StockIn_SalesReturn
        StockIn_Adjust
        StockOut_Sales
        StockOut_PurchaseReturn
        StockOut_Adjust
        StockIn_Process
        StockIn_BackProcess
        StockOut_Process
    End Enum

    Public Function GetJmlByAsset(ByVal vPrimaryAsset As Double) As Double
        Dim DTA As New DseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlByAsset(vPrimaryAsset).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlByAsset(vPrimaryAsset).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByAssetNInOut(ByVal vPrimaryAsset As Double, ByVal vInOut As Integer) As Double
        Dim DTA As New DseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlByAssetNInOut(vPrimaryAsset, vInOut).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlByAssetNInOut(vPrimaryAsset, vInOut).Value
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
