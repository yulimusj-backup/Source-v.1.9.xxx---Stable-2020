Partial Class DseMasterProduct_BOM1
    Partial Class DataTable2DataTable

        Private Sub DataTable2DataTable_ColumnChanging(ByVal sender As System.Object, ByVal e As System.Data.DataColumnChangeEventArgs) Handles Me.ColumnChanging
            If (e.Column.ColumnName = Me.Primary_RMColumn.ColumnName) Then
                'Add user code here
            End If

        End Sub

    End Class

End Class
