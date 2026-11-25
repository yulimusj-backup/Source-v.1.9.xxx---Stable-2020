Option Explicit On
Option Strict Off

Public Class FrmTransaksiStockTransfer_BatasanMinus
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public TblMinus As New dseTransaksiStockTransfer.tbl_batasan_minusDataTable
    Dim rowMinus As dseTransaksiStockTransfer.tbl_batasan_minusRow
    Dim ObjCD As New Core_Module.ClsCekData

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Me.Close()
    End Sub

    Private Sub FrmTransaksiStockTransfer_BatasanMinus_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim rd As DataGridViewRow

        Try
            Me.ColQty.HeaderText = "Minus Qty"
            Me.Text = "The following product(s) will be minus. Process cancelled."
            Me.DgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            TblMinus.DefaultView.RowFilter = ""
            TblMinus.DefaultView.Sort = "[Product Code], [batch], [Location]"

            Me.DgvItem.Rows.Clear()
            Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each dv As Data.DataRowView In TblMinus.DefaultView
                rowMinus = dv.Row
                Me.DgvItem.Rows.Add()
                rd = Me.DgvItem.Rows(Me.DgvItem.RowCount - 1)
                rd.Cells(Me.ColNo.Index).Value = Me.DgvItem.RowCount
                rd.Cells(Me.ColProductCode.Index).Value = rowMinus.Product_Code
                rd.Cells(Me.ColBatchNo.Index).Value = rowMinus.batch
                rd.Cells(Me.ColLocation.Index).Value = rowMinus.Location
                rd.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(rowMinus.qty, rowMinus._Decimal)
                rd.Cells(Me.ColUnit.Index).Value = rowMinus.Unit_Code
            Next
            Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvItem.CellBeginEdit
        e.Cancel = True
    End Sub

    Private Sub DgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvItem.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class