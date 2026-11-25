Partial Class DseTransaksiPackingList
    Partial Class tbltrans_packinglist_view0DataTable

    End Class

    Partial Class DataTable21DataTable

    End Class

    Partial Class tbltrans_packinglist_view1DataTable

        Private Sub tbltrans_packinglist_view1DataTable_ColumnChanging(ByVal sender As System.Object, ByVal e As System.Data.DataColumnChangeEventArgs) Handles Me.ColumnChanging
            If (e.Column.ColumnName = Me.Order_BalanceColumn.ColumnName) Then
                'Add user code here
            End If

        End Sub

    End Class

End Class
