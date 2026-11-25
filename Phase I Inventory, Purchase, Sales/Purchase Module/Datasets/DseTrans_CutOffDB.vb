Partial Class DseTrans_CutOffDB

    Partial Class tbltrans_purchasereturn1DataTable

        Private Sub tbltrans_purchasereturn1DataTable_ColumnChanging(ByVal sender As System.Object, ByVal e As System.Data.DataColumnChangeEventArgs) Handles Me.ColumnChanging
            If (e.Column.ColumnName = Me.PrimaryColumn.ColumnName) Then
                'Add user code here
            End If

        End Sub

    End Class

End Class
