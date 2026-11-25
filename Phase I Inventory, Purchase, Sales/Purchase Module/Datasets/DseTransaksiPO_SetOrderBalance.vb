Partial Class DseTransaksiPO_SetOrderBalance
    Partial Class tbltrans_packinglist_itemDataTable

        Private Sub tbltrans_packinglist_itemDataTable_tbltrans_packinglist_itemRowChanging(ByVal sender As System.Object, ByVal e As tbltrans_packinglist_itemRowChangeEvent) Handles Me.tbltrans_packinglist_itemRowChanging

        End Sub

        Private Sub tbltrans_packinglist_itemDataTable_ColumnChanging(ByVal sender As System.Object, ByVal e As System.Data.DataColumnChangeEventArgs) Handles Me.ColumnChanging
            If (e.Column.ColumnName = Me.Total_QtyColumn.ColumnName) Then
                'Add user code here
            End If

        End Sub

    End Class

End Class
