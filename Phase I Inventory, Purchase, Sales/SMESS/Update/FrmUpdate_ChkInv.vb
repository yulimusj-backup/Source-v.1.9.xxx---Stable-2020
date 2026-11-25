Public Class FrmUpdate_ChkInv
    Public TBLInv As DseUpdate_v2018.tbltrans_salesinvoiceDataTable
    Dim rowInv As DseUpdate_v2018.tbltrans_salesinvoiceRow

    Private Sub FrmUpdate_ChkInv_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim rowDGV As DataGridViewRow

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            For Each rowInv In TBLInv.Rows
                Me.dgvInv.Rows.Add()
                rowDGV = Me.dgvInv.Rows(Me.dgvInv.Rows.Count - 1)
                rowDGV.Cells(Me.ColInvNo.Index).Value = rowInv.No_Transaksi
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class