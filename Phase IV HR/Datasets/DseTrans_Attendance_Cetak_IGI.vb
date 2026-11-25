Partial Class DseTrans_Attendance_Cetak_IGI
  

    Partial Class DataTable11DataTable

        Private Sub DataTable11DataTable_ColumnChanging(ByVal sender As System.Object, ByVal e As System.Data.DataColumnChangeEventArgs) Handles Me.ColumnChanging
            If (e.Column.ColumnName = Me.Total_PermitColumn.ColumnName) Then
                'Add user code here
            End If

        End Sub

    End Class

End Class
