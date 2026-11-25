Option Explicit On
Option Strict Off

Public Class ClsMasterProduct_Verifikasi
    Private mVerified As Integer = EnumVerified.No

    Public Enum EnumVerified As Integer
        No
        Yes
    End Enum

    Public Function SimpanPerubahanVerify(ByVal vVerified As EnumVerified, ByVal vUserCode As String, ByVal vPrimary As Double) As Boolean
        Dim DTA As New dseMasterProductTableAdapters.tblmaster_product4TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.UpdateQuery(vVerified, vUserCode, Now, vPrimary) > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
