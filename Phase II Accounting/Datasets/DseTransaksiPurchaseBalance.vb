Partial Class DseTransaksiPurchaseBalance
    Partial Class ListOfCurrencyDataTable

        Private Sub ListOfCurrencyDataTable_ColumnChanging(ByVal sender As System.Object, ByVal e As System.Data.DataColumnChangeEventArgs) Handles Me.ColumnChanging
            If (e.Column.ColumnName = Me.Balance_OriginalColumn.ColumnName) Then
                'Add user code here
            End If

        End Sub

    End Class

    Partial Class APAR_PrintDataTable

    End Class

    Partial Class tbltrans_historicalpurchasebalancemanualDataTable


    End Class

End Class
