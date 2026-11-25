Partial Class DseTrans_Gaji_MasaSblm
    Partial Class tbltrans_salary_prev_incomeDataTable

        Private Sub tbltrans_salary_prev_incomeDataTable_ColumnChanging(ByVal sender As System.Object, ByVal e As System.Data.DataColumnChangeEventArgs) Handles Me.ColumnChanging
            If (e.Column.ColumnName = Me.nett_incomeColumn.ColumnName) Then
                'Add user code here
            End If

        End Sub

    End Class

End Class
