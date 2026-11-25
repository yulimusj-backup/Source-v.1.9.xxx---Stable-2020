Partial Class DseTransaksiSO
    Partial Class DataTable5DataTable

        Private Sub DataTable5DataTable_ColumnChanging(ByVal sender As System.Object, ByVal e As System.Data.DataColumnChangeEventArgs) Handles Me.ColumnChanging
            If (e.Column.ColumnName = Me.Primary_UnitColumn.ColumnName) Then
                'Add user code here
            End If

        End Sub

    End Class

End Class
