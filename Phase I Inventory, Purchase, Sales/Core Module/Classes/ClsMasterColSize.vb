Option Explicit On
Option Strict Off

Public Class ClsMasterColSize
    Public MaxWaitTime As Double = 3000 'ms
    Public Enum EnumTipeTransaksi As Integer
        MasterProduct
        MasterProduct1
    End Enum

    Public Sub SimpanTBL(ByVal vTBL As DseMasterColSize.tblmaster_columns_sizeDataTable)
        Dim DTA As New DseMasterColSizeTableAdapters.tblmaster_columns_sizeTableAdapter
        Dim row As DseMasterColSize.tblmaster_columns_sizeRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vTBL.Rows.Count > 0 Then
                row = vTBL.Rows(0)
                DTA.Delete(row.tipe_transaksi, row.user)
                For Each row In vTBL.Rows
                    DTA.Insert(row.tipe_transaksi, row.user, row.col_index, row.col_size)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsiTBL(ByVal vUserCode As String, ByVal vTipeTransaksi As EnumTipeTransaksi, ByVal vTBL As DseMasterColSize.tblmaster_columns_sizeDataTable)
        Dim DTA As New DseMasterColSizeTableAdapters.tblmaster_columns_sizeTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy1(vTBL, vTipeTransaksi, vUserCode)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub
End Class
