Public Class FrmUpdate_CheckData_ItemDO
    Public TblRef As New DseUpdate_CheckData.tbldo_itemDataTable
    Dim rowRef As DseUpdate_CheckData.tbldo_itemRow

    Private Sub FrmUpdate_CheckData_ItemDO_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowRef In Me.TblRef.Rows
                Me.dgvHasil.Rows.Add()
                rowDGV = Me.dgvHasil.Rows(Me.dgvHasil.Rows.Count - 1)
                rowDGV.Cells(Me.ColCodeNo.Index).Value = rowRef.Product_Code
                rowDGV.Cells(Me.ColDONo.Index).Value = rowRef.No_Transaksi
                rowDGV.Cells(Me.ColNo.Index).Value = rowDGV.Index + 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class