Option Explicit On
Option Strict Off

'copy dari ClsTransaksiBalanceManual_DP dari modul akunting
Public Class ClsTransaksiBalanceDP
    Public Enum EnumTipeDP As Integer
        DP_AOP
        DP_Invoice
    End Enum

    Public Function GetJml(ByVal vTipeDP As EnumTipeDP, ByVal vPrimaryDocAcuan As Double) As Double
        Dim DTA As New DseBalanceDPTableAdapters.tbltrans_historicalbalancemanual_dpTableAdapter
        Dim RetVal As Nullable(Of Decimal)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetDPAmount(CInt(vTipeDP), vPrimaryDocAcuan)

            If RetVal.HasValue = True Then
                Return RetVal.Value
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
