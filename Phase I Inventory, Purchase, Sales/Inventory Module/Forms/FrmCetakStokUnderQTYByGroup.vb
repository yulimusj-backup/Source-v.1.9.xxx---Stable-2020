Option Explicit On
Option Strict Off

Public Class FrmCetakStockUnderQTYByGroup
    Public primary As Double = -1
    Private Sub FrmCetakStokUnderQTYByGroup_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim report As New CRPStockUnderQTYByGroup
        Dim dta As New dsLaporanStockUnderQTYByGroupTableAdapters.TblStockMinimumTableAdapter
        Dim tbl As dsLaporanStockUnderQTYByGroup.TblStockMinimumDataTable
        Dim Row As dsLaporanStockUnderQTYByGroup.TblStockMinimumRow
        Dim ds As New dsLaporanStockUnderQTYByGroup

        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim rowlaporan As dsLaporanStockUnderQTYByGroup.tblcetakRow

        Try
            tbl = dta.GetDataBy_PrimaryProduct(primary)
            For Each Row In tbl.Rows
                rowlaporan = ds.tblcetak.NewRow()
                rowlaporan.CategoryNDescCode = Row.Category_Code + "-" + Row.Desc_Code
                rowlaporan.Code_No = Row.Category_Code
                rowlaporan.Description = Row.Description
                rowlaporan.Unit = Row.Unit_Code
                rowlaporan.Min_Qty = Row.Minimum_Stock
                rowlaporan.Qty = Row.Jumlah_Qty_Tbl_Product_Qty
                ds.tblcetak.Rows.Add(rowlaporan)
            Next
            report.SetDataSource(ds)
            Me.CrystalReportViewer1.ReportSource = report
            Me.CrystalReportViewer1.Show()
            Windows.Forms.Cursor.Current = Cursors.Arrow

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
            ds.Dispose()
        End Try
    End Sub

    Private Sub BtnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Me.Close()
    End Sub
End Class