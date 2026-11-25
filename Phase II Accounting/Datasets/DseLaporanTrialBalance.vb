

Partial Public Class DseLaporanTrialBalance

    Partial Class Trial_BalanceDataTable

        Private Sub Trial_BalanceDataTable_ColumnChanging(ByVal sender As System.Object, ByVal e As System.Data.DataColumnChangeEventArgs) Handles Me.ColumnChanging
            If (e.Column.ColumnName = Me.Primary_AkunColumn.ColumnName) Then
                'Add user code here
            End If

        End Sub

    End Class

    Partial Class TblCetakDataTable

        Private Sub TblCetakDataTable_ColumnChanging(ByVal sender As System.Object, ByVal e As System.Data.DataColumnChangeEventArgs) Handles Me.ColumnChanging


        End Sub

    End Class

End Class
