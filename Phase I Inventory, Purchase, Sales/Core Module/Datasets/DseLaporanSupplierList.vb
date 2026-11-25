

Partial Public Class DseLaporanSupplierList
    Partial Class tblItem_CPDataTable

        Private Sub tblItem_CPDataTable_tblItem_CPRowChanging(ByVal sender As System.Object, ByVal e As tblItem_CPRowChangeEvent) Handles Me.tblItem_CPRowChanging

        End Sub

    End Class

    Partial Class tbl_cetakDataTable

        Private Sub tbl_cetakDataTable_ColumnChanging(ByVal sender As System.Object, ByVal e As System.Data.DataColumnChangeEventArgs) Handles Me.ColumnChanging
            'If (e.Column.ColumnName = Me.Primary_Item_AddressColumn.ColumnName) Then                'Add user code here
            'End If

        End Sub

    End Class

End Class
