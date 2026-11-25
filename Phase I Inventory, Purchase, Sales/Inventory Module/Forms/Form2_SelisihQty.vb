Option Explicit On
Option Strict Off

Public Class Form2_SelisihQty
    Public TBLRef As DataSet2_CekSelisihQty.DataTable1DataTable
    Dim rowRef As DataSet2_CekSelisihQty.DataTable1Row
    Dim ObjCD As New Core_Module.ClsCekData

    Private Sub Form2_SelisihQty_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim rowDGV As DataGridViewRow

        Try
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            For Each rowRef In Me.TBLRef.Rows
                Me.dgvItem.Rows.Add()
                rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)

                rowDGV.Cells(Me.ColCodeNo.Index).Value = rowRef.Product_Code
                rowDGV.Cells(Me.ColDesc.Index).Value = rowRef.Description
                rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvItem.Rows.Count
                rowDGV.Cells(Me.ColQtyClosing.Index).Value = Me.ObjCD.DigitGrouping(rowRef.Ending_Qty, rowRef._Decimal)
                rowDGV.Cells(Me.ColQtySC.Index).Value = Me.ObjCD.DigitGrouping(rowRef.Qty_SC, rowRef._Decimal)
                rowDGV.Cells(Me.ColUnit.Index).Value = rowRef.Unit_Code
            Next

            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class