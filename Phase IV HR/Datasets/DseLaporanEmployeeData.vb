

Partial Public Class DseLaporanEmployeeData
    Partial Class Format_1DataTable

        Private Sub Format_1DataTable_ColumnChanging(ByVal sender As System.Object, ByVal e As System.Data.DataColumnChangeEventArgs) Handles Me.ColumnChanging
            If (e.Column.ColumnName = Me.PrimaryColumn.ColumnName) Then
                'Add user code here
            End If

        End Sub

    End Class

    Partial Class Format_2DataTable

    End Class

End Class
