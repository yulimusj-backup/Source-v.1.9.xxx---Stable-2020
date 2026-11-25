Option Strict Off
Option Explicit On

Public Class FrmCetakStockByBatchNumberByGroup
    Public primary As Double = -1
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
    Private Sub FrmLaporanStockByBatchNumberByGroup_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim report As New CRPStockByBatchNumberByGroup
        Dim dta As New DsLaporanStockByBatchNumberByGroupTableAdapters.DataCetakTableAdapter
        Dim tbl As DsLaporanStockByBatchNumberByGroup.DataCetakDataTable
        Dim Row As DsLaporanStockByBatchNumberByGroup.DataCetakRow
        Dim ds As New DsLaporanStockByBatchNumberByGroup

        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim rowlaporan As DsLaporanStockByBatchNumberByGroup.TblCetakRow

        Try
            tbl = dta.GetData(primary)
            For Each Row In tbl.Rows
                rowlaporan = ds.TblCetak.NewRow()
                rowlaporan.Category_CodeNDesc_Code = Row.Category_Code + "-" + Row.Desc_Code
                rowlaporan.Code_Number = Row.Code_Number
                rowlaporan.Description = Row.Description
                rowlaporan.Unit = Row.Unit
                rowlaporan.Batch_Number = Row.Batch_Number
                rowlaporan.Qty_Per_Batch = Row.Qty_Per_Batch
                ds.TblCetak.Rows.Add(rowlaporan)
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
End Class