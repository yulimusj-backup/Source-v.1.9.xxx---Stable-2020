Option Explicit On
Option Strict Off

Public Class FrmCetakStockByBatchNumber
    Public primary As Double = -1
    Private Sub CetakSemua()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim ds As New dsLaporanStockByBatchNumber
        Dim report As New CRPStockByBatchNumber
        Dim DTA As New dsLaporanStockByBatchNumberTableAdapters.TblCetakTableAdapter
        Try
            DTA.FillBy(ds.TblCetak, primary)
            Dim baris As Integer = ds.TblCetak.Rows.Count
            report.SetDataSource(ds)
            Me.CrystalReportViewer1.ReportSource = report
            Me.CrystalReportViewer1.Show()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ds.Dispose()
        End Try
    End Sub

    Private Sub FrmCetakStockByBatchNumber_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.CetakSemua()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Me.Close()
    End Sub
End Class