Partial Class DseUpdate_v180
    Partial Class tbltrans_do_itemDataTable

        Private Sub tbltrans_do_itemDataTable_tbltrans_do_itemRowChanging(ByVal sender As System.Object, ByVal e As tbltrans_do_itemRowChangeEvent) Handles Me.tbltrans_do_itemRowChanging

        End Sub

    End Class

    Partial Class tbltrans_historicalpurchasereturnbalancemanualDataTable

        Private Sub tbltrans_historicalpurchasereturnbalancemanualDataTable_ColumnChanging(ByVal sender As System.Object, ByVal e As System.Data.DataColumnChangeEventArgs) Handles Me.ColumnChanging
            If (e.Column.ColumnName = Me.PrimaryColumn.ColumnName) Then
                'Add user code here
            End If

        End Sub

    End Class

End Class
