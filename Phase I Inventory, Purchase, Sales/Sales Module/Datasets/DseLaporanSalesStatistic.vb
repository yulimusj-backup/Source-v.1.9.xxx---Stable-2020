Partial Class DseLaporanSalesStatistic
    Partial Class Group_By_Cust_Group_CodeDataTable

        Private Sub Group_By_Cust_Group_CodeDataTable_ColumnChanging(ByVal sender As System.Object, ByVal e As System.Data.DataColumnChangeEventArgs) Handles Me.ColumnChanging
            If (e.Column.ColumnName = Me.JanuaryColumn.ColumnName) Then
                'Add user code here
            End If

        End Sub

    End Class

End Class
