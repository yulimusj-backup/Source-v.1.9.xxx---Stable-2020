

Partial Public Class Dse2

    Partial Class DataTable3DataTable

        Private Sub DataTable3DataTable_ColumnChanging(ByVal sender As System.Object, ByVal e As System.Data.DataColumnChangeEventArgs) Handles Me.ColumnChanging
            If (e.Column.ColumnName = Me.Primary_WarehouseColumn.ColumnName) Then
                'Add user code here
            End If

        End Sub

    End Class

End Class
