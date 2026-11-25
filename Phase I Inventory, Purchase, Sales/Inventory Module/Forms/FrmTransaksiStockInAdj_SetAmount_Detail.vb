Option Explicit On
Option Strict Off
Public Class FrmTransaksiStockInAdj_SetAmount_Detail

    Public vPrimaryProduct As Double = -1
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public vDecimalDefault As Integer = 2

    Dim objCD As New Core_Module.ClsCekData
    Dim Dta As New dseTransaksiStockInOut_SetAmountTableAdapters.DataTable1TableAdapter
    Dim Tbl As New dseTransaksiStockInOut_SetAmount.DataTable1DataTable
    Dim Row As dseTransaksiStockInOut_SetAmount.DataTable1Row

    Private Sub FrmTransaksiStockInAdj_SetAmount_Detail_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.Cursor = Cursors.WaitCursor
            Me.dgvDetail.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Dta.FillBy_primary_prod(Tbl, Me.vPrimaryProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiStockInAdj_SetAmount_Detail_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Dim RowDgv As DataGridViewRow
        Try
            Me.dgvDetail.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvDetail.Rows.Clear()
            For Each Row In Tbl
                RowDgv = Me.dgvDetail.Rows(Me.dgvDetail.Rows.Add)
                RowDgv.Cells(Me.ColDetail_Amount.Index).Value = Me.objCD.DigitGrouping(Row.Amount, Me.vDecimalDefault, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)
                RowDgv.Cells(Me.ColDetail_Date.Index).Value = Format(Row.Tanggal_Transaksi, "dd MMM yyyy")
                RowDgv.Cells(Me.ColDetail_DocNo.Index).Value = Row.No_Transaksi
                RowDgv.Cells(Me.ColDetail_PrimaryProd.Index).Value = Row.Primary_Product
                RowDgv.Cells(Me.ColDetail_PrimaryTrans.Index).Value = Row.Primary_Transaksi
                RowDgv.Cells(Me.ColDetail_Qty.Index).Value = Me.objCD.DigitGrouping(Row.Qty, Row._Decimal, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)
                RowDgv.Cells(Me.ColDetail_Unit.Index).Value = Row.Unit_Code
                RowDgv.Cells(Me.ColDetail_UnitPrice.Index).Value = Me.objCD.DigitGrouping(Row.Unit_Price, Me.vDecimalDefault, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)
                RowDgv.Cells(Me.ColDetail_No.Index).Value = Me.dgvDetail.RowCount.ToString
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.dgvDetail.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.Cursor = Cursors.Default
        End Try
    End Sub 

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
End Class