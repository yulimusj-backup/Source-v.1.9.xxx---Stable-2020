Partial Class dseLaporanSalesOrderBalance2
    Partial Class TblCetakDataTable

        Private Sub TblCetakDataTable_ColumnChanging(ByVal sender As System.Object, ByVal e As System.Data.DataColumnChangeEventArgs) Handles Me.ColumnChanging
            If (e.Column.ColumnName = Me.CustomerCodeColumn.ColumnName) Then
                'Add user code here
            End If

        End Sub

    End Class

    Partial Class laporan_sales_orderbalance_view9bDataTable

        Private Sub laporan_sales_orderbalance_view9bDataTable_ColumnChanging(ByVal sender As System.Object, ByVal e As System.Data.DataColumnChangeEventArgs) Handles Me.ColumnChanging
            If (e.Column.ColumnName = Me.Primary_SOColumn.ColumnName) Then
                'Add user code here
            End If

        End Sub

    End Class

    Partial Class DataTable1DataTable

    End Class

End Class
