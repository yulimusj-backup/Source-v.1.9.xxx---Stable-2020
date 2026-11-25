Partial Class DseTransaksiQuotation
    Partial Class TblCetakIsiDataTable

        Private Sub TblCetakIsiDataTable_ColumnChanging(ByVal sender As System.Object, ByVal e As System.Data.DataColumnChangeEventArgs) Handles Me.ColumnChanging
            If (e.Column.ColumnName = Me.Unit_CodeColumn.ColumnName) Then
                'Add user code here
            End If

        End Sub

    End Class

End Class
