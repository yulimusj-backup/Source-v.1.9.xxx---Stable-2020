Partial Class DseTrans_Export_eSPT


    Partial Class tbltrans_salary_yearlyDataTable

        Private Sub tbltrans_salary_yearlyDataTable_ColumnChanging(ByVal sender As System.Object, ByVal e As System.Data.DataColumnChangeEventArgs) Handles Me.ColumnChanging
            If (e.Column.ColumnName = Me.Masa_PajakColumn.ColumnName) Then
                'Add user code here
            End If

        End Sub

    End Class

End Class
