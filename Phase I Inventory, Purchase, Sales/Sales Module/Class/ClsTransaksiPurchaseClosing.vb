Option Explicit On
Option Strict Off

Public Class ClsTransaksiPurchaseClosing
    Public Function GetPurchaseClosingMinimumDate() As Date
        Dim DTA As New DseTransaksiPurchaseClosingTableAdapters.DataTable1TableAdapter
        Dim TBL As New DseTransaksiPurchaseClosing.DataTable1DataTable
        Dim row As DseTransaksiPurchaseClosing.DataTable1Row

        Try
            If DTA.Fill(TBL) > 0 Then
                row = TBL.Rows(0)
                If row.IsMinimum_Closing_PeriodNull = True Then
                    Return #1/1/1950#
                Else
                    Return row.Minimum_Closing_Period
                End If
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetLastClosingPeriod() As Date
        Dim DTA As New DseTransaksiPurchaseClosingTableAdapters.DataTable2TableAdapter
        Dim TBL As New DseTransaksiPurchaseClosing.DataTable2DataTable
        Dim row As DseTransaksiPurchaseClosing.DataTable2Row

        Try
            If DTA.Fill(TBL) > 0 Then
                row = TBL.Rows(0)
                If row.IsMaxClosingPeriodNull = True Then
                    Return #1/1/1950#
                End If
                Return row.MaxClosingPeriod
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function
End Class
