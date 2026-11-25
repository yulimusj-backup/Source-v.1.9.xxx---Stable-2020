Partial Public Class DseTrans_SPL
    Partial Class tbltrans_spl_overtime_order_approval_settingDataTable

        Private Sub tbltrans_spl_overtime_order_approval_settingDataTable_ColumnChanging(ByVal sender As System.Object, ByVal e As System.Data.DataColumnChangeEventArgs) Handles Me.ColumnChanging
            If (e.Column.ColumnName = Me.PrimaryColumn.ColumnName) Then
                'Add user code here
            End If

        End Sub

    End Class

End Class
