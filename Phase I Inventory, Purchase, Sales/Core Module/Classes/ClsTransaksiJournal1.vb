Option Explicit On
Option Strict Off

Public Class ClsTransaksiJournal1
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

    Public Enum EnumTipeJurnal2 As Integer
        Jurnal_NoGrouping
        GeneralJurnal_MonthlyRevaluation
        Cash_n_Bank
        AssetAcquistion_Other
        AssetRevaluation_Other
        AssetDepreciation
        AssetDisposal
        AssetSetFinish
    End Enum

    Public Function GetJmlDepAsset(ByVal vPrimaryAsset As Double) As Double
        Dim DTA As New DseMasterAssetTableAdapters.tbltrans_jurnal_item_depresiasi_asetTableAdapter
        Dim RetVal As Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetJmlByPrimaryAsset(vPrimaryAsset).Value
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByPrimaryAsset(ByVal vPrimaryAsset As Double) As Double
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQueryGetJmlByPrimaryAsset(vPrimaryAsset).HasValue = True Then
                Return DTA.ScalarQueryGetJmlByPrimaryAsset(vPrimaryAsset).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByTipeNPrimaryAsset(ByVal vPrimaryAsset As Double, ByVal tipe As Integer) As Double
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQueryGetJmlByTipeNPrimaryAsset(vPrimaryAsset, tipe).HasValue = True Then
                Return DTA.ScalarQueryGetJmlByTipeNPrimaryAsset(vPrimaryAsset, tipe).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTipeJournal2(ByVal vPrimaryAsset As Double) As Integer
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetTipeJournal2(vPrimaryAsset).HasValue = True Then
                Return DTA.ScalarQuery_GetTipeJournal2(vPrimaryAsset).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetClosingDate() As Date
        Dim DTA As New Core_Module.DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetClosingDate.HasValue = True Then
                Return DTA.ScalarQuery_GetClosingDate.Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
