Option Explicit On
Option Strict Off
Public Class ClsDBVersion
    Public Sub SetDBVersion(ByVal vNewVersion As String)
        Dim DTA As New DseDBVersionTableAdapters.tbltrans_dbversionTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Insert(vNewVersion, Today, "")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function GetDBVersion() As String
        Dim DTA As New DseDBVersionTableAdapters.tbltrans_dbversionTableAdapter
        Dim TBL As New DseDBVersion.tbltrans_dbversionDataTable
        Dim row As DseDBVersion.tbltrans_dbversionRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(TBL.Rows.Count - 1)
                Return row.version_number
            End If

            Return ""
        Catch ex As Exception
            'tidak menggunakan pesan error yang biasa karena untuk memfasilitasi database yang sebelumnya tidak menggunakan fitur ini.
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function
End Class
