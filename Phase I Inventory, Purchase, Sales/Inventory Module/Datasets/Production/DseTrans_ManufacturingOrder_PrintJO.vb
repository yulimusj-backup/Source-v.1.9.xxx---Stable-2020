Partial Class DseTrans_ManufacturingOrder_PrintJO
    Partial Class tblCetakBahanDataTable

        Private Sub tblCetakBahanDataTable_ColumnChanging(ByVal sender As System.Object, ByVal e As System.Data.DataColumnChangeEventArgs) Handles Me.ColumnChanging

        End Sub
    End Class

    Partial Class tblCetakDataTable

        Private Sub tblCetakDataTable_ColumnChanging(ByVal sender As System.Object, ByVal e As System.Data.DataColumnChangeEventArgs) Handles Me.ColumnChanging
            If (e.Column.ColumnName = Me._Sche_R_PColumn.ColumnName) Then
                'Add user code here
            End If

        End Sub

    End Class

    Partial Class tblCetakRPBahanDataTable

    End Class

End Class
