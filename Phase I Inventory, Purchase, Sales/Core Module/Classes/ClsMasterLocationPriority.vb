Public Class ClsMasterLocationPriority
    Public Enum EnumTipeTransaksi As Integer
        StockOutAdjust
        StockOutSales
        StockBundling
    End Enum

    Public Sub GetLocationPrioriry(ByVal vTBL As Core_Module.DseMasterLocationPriority.tblmaster_location_priority1_DataTable, _
    ByVal vUserCode As String, ByVal vTipeTransaksi As EnumTipeTransaksi)
        Dim DTALocation As New Core_Module.DseMasterLocationPriorityTableAdapters.tblmaster_location_priority1_TableAdapter

        Try
            DTALocation.Connection.ConnectionString = Core_Module.GetConstring
            vTBL.Rows.Clear()
            DTALocation.Fill(vTBL, vUserCode, CInt(vTipeTransaksi))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTALocation.Dispose()
        End Try
    End Sub
End Class
