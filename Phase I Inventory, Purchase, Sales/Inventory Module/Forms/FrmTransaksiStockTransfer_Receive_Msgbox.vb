Option Explicit On
Option Strict Off

Public Class FrmTransaksiStockTransfer_Receive_Msgbox
    Public TblReason As dseTransaksiStockTransfer.tbl_cancel_receivedDataTable
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Private Sub FrmTransaksiStockTransfer_Receive_Msgbox_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim row As dseTransaksiStockTransfer.tbl_cancel_receivedRow
        Dim rd As DataGridViewRow

        Try
            Me.MinimumSize = Me.Size
            Me.MaximumSize = Me.Size
            Me.DgvItem.ReadOnly = True
            Me.DgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvItem.Rows.Clear()
            Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each row In TblReason
                Me.DgvItem.Rows.Add()
                rd = Me.DgvItem.Rows(Me.DgvItem.RowCount - 1)
                rd.Cells(Me.ColNo.Index).Value = rd.Index + 1
                rd.Cells(Me.ColReceivedDocNo.Index).Value = row.No_Doc
                rd.Cells(Me.ColReason.Index).Value = row.Reason
            Next
            Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Me.Label1.Text = "Total Rows : " & Me.DgvItem.Rows.Count
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Me.Close()
    End Sub
End Class