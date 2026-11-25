Option Explicit On
Option Strict Off
Public Class ClsDBVersion
    Public Function GetDBVersion() As String
        Dim DTA As New DseDBVersionTableAdapters.tbltrans_dbversionTableAdapter
        Dim TBL As New DseDBVersion.tbltrans_dbversionDataTable
        Dim row As DseDBVersion.tbltrans_dbversionRow

        Try
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
