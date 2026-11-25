Option Explicit On
Option Strict Off

Public Class ClsTransaksiAOPItem
    Public Function GetNoTransaksi(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter
        Dim Retval As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            Retval = DTA.ScalarQuery_GetNoTransaksiAOP(vPrimary)
            If String.IsNullOrEmpty(Retval) = True Then
                Return ""
            Else
                Return Retval
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''dipakai untuk Class GRN
    Public Function GetJmlByGRN(ByVal PrimaryGRN As Double) As Integer
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aop_itemTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetJmlByPrimaryGRN(PrimaryGRN).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiTBLGRN(ByVal vPrimaryAOP As Double, ByVal TBL As DseTransaksiAOP.DataTable9DataTable) As Boolean
        Dim DTA As New DseTransaksiAOPTableAdapters.DataTable9TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimaryAOP)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
