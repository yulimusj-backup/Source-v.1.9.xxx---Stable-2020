Option Explicit On
Option Strict Off

Public Class FrmTrans_CutOffDB_ExceptionQty
    Public TblShow As Core_Module.DseTrans_CutOffDB_Core.Tbl_PerbandinganQtyDataTable
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim ObjCD As New Core_Module.ClsCekData

    Private Sub FrmTrans_CutOffDB_ExceptionQty_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim rowShow As Core_Module.DseTrans_CutOffDB_Core.Tbl_PerbandinganQtyRow
        Dim rowD As DataGridViewRow

        Try
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.DgvQty.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvQty.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DgvQty.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            For Each rowShow In Me.TblShow
                Me.DgvQty.Rows.Add()
                rowD = Me.DgvQty.Rows(Me.DgvQty.Rows.Count - 1)
                rowD.Cells(Me.ColProdCode.Index).Value = rowShow.Product_Code
                rowD.Cells(Me.ColProdQty.Index).Value = Me.ObjCD.BuangNolDibelakangDecimal(rowShow.Total_Prod_Qty)
                rowD.Cells(Me.ColProdQty_StockCard.Index).Value = Me.ObjCD.BuangNolDibelakangDecimal(rowShow.Total_SC_Qty)
            Next
            Me.DgvQty.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.DgvQty.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class