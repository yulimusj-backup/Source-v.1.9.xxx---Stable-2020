

Partial Public Class DseLaporanPurchaseReturnDetail
    Partial Class Total_CurrDataTable

    



    End Class

    Partial Class Tbl_CetakDataTable

        Private Sub Tbl_CetakDataTable_Tbl_CetakRowChanging(ByVal sender As System.Object, ByVal e As Tbl_CetakRowChangeEvent) Handles Me.Tbl_CetakRowChanging

        End Sub

        Private Sub Tbl_CetakDataTable_ColumnChanging(ByVal sender As System.Object, ByVal e As System.Data.DataColumnChangeEventArgs) Handles Me.ColumnChanging
            If (e.Column.ColumnName = Me.VATColumn.ColumnName) Then
                'Add user code here
            End If

        End Sub

    End Class

End Class
