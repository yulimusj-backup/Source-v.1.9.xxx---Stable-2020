Option Explicit On
Option Strict Off

Public Class FrmLaporanStockTaking_FinalCheck
    Public DS As dseLaporanStockTaking

    Dim vReport As New CrpLaporanStockTaking_DataCheck

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub FrmLaporanStockTaking_FinalCheck_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'tampilkan report
            vReport.SetDataSource(DS)

            vcrv.ReportSource = vReport
            vcrv.ShowRefreshButton = False
            vcrv.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class