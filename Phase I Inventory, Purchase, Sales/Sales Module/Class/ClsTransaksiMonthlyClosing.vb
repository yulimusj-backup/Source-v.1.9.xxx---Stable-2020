Public Class ClsTransaksiMonthlyClosing
    Public Enum EnumTipeItemDetail As Integer
        BeginingBalanceItem
        InItem
        OutItem
        EndingBalanceItem
    End Enum

    Public Enum EnumTipeTranskasiAsal As Integer
        InPurchase '0
        InSR '1
        InAdjust '2
        InStockTakeAdjust '3
        OutBundling '4
        OutSales '5
        OutPR '6
        OutStockTakeAdjust '7
        OutAdjust '8
        InBundling '9
        OverheadCost '10
    End Enum

    Public Function GetPrimaryClosing(ByVal vPeriod As Date) As Double
        Dim DTA As New DseTransaksiPurchaseClosingTableAdapters.DataTable1TableAdapter
        Dim Dummy As Nullable(Of Double)

        Try
            Dummy = DTA.ScalarQuery_GetPrimaryClosing(vPeriod)
            If Dummy.HasValue = True Then
                Return Dummy.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
